"""
Parser module for extracting key-value pairs from text using Google Gemini LLM
Follows the strategy from Final Plan Key Value Comments.txt
"""

import re
import json
from typing import Dict, List, Optional, Any, Tuple
from google import genai
from google.genai import types
from .logger import logger


def parse_key_value_pairs(
    text: str, 
    api_key: str,
    custom_keys: Optional[List[str]] = None,
    extraction_method: str = "Gemini LLM",
    doc_id: str = "D1"
) -> List[Dict[str, Any]]:
    """
    Parse text and extract key-value pairs using Google Gemini LLM
    Returns list of dicts with keys: key, value, comments
    
    Args:
        text: Input text to parse
        api_key: Google Gemini API key
        custom_keys: Optional list of specific keys to prioritize
        extraction_method: Method to use for extraction (default: Gemini LLM)
        doc_id: Document identifier (e.g., D1, D2)
        
    Returns:
        list: List of extracted records with key, value, comments fields
    """
    if not api_key:
        raise ValueError("Google Gemini API key is required")
    
    # Preprocess text into chunks with paragraph tracking
    chunks = preprocess_and_chunk(text, doc_id)
    
    # Extract from each chunk using Gemini
    all_extractions = []
    for chunk_data in chunks:
        chunk_text = chunk_data['text']
        doc_id_chunk = chunk_data['doc_id']
        paragraph_index = chunk_data['paragraph_index']
        char_offset = chunk_data['offset']
        
        extracted = extract_with_gemini(chunk_text, doc_id_chunk, paragraph_index, char_offset, custom_keys, api_key)
        all_extractions.extend(extracted)
    
    # Convert to final 3-column format (key, value, comments)
    final_rows = convert_to_three_columns(all_extractions)
    
    return final_rows


def preprocess_and_chunk(text: str, doc_id: str = "D1", chunk_size: int = 2000, overlap: int = 200) -> List[Dict[str, Any]]:
    """
    Preprocess text and split into overlapping chunks with metadata
    Tracks paragraph indices and character offsets per Final Plan requirements
    
    Args:
        text: Input text
        doc_id: Document identifier
        chunk_size: Approximate chunk size in characters
        overlap: Overlap between chunks
        
    Returns:
        list: List of chunk dictionaries with text, doc_id, paragraph_index, offset
    """
    # Normalize text
    text = text.strip()
    
    # Split into paragraphs and track indices
    paragraphs = [p.strip() for p in text.split('\n\n') if p.strip()]
    
    chunks = []
    current_chunk = []
    current_length = 0
    char_offset = 0
    paragraph_start_index = 0
    
    for para_idx, para in enumerate(paragraphs):
        para_len = len(para)
        
        if current_length + para_len > chunk_size and current_chunk:
            # Create chunk
            chunk_text = '\n\n'.join(current_chunk)
            chunks.append({
                'text': chunk_text,
                'doc_id': doc_id,
                'paragraph_index': paragraph_start_index,
                'offset': char_offset
            })
            
            # Keep last paragraph for overlap
            if len(current_chunk) > 1:
                current_chunk = [current_chunk[-1]]
                current_length = len(current_chunk[0])
                paragraph_start_index = para_idx - 1
            else:
                current_chunk = []
                current_length = 0
                paragraph_start_index = para_idx
        
        current_chunk.append(para)
        current_length += para_len
        char_offset += para_len + 2  # +2 for \n\n
    
    # Add remaining chunk
    if current_chunk:
        chunk_text = '\n\n'.join(current_chunk)
        chunks.append({
            'text': chunk_text,
            'doc_id': doc_id,
            'paragraph_index': paragraph_start_index,
            'offset': char_offset
        })
    
    return chunks if chunks else [{'text': text, 'doc_id': doc_id, 'paragraph_index': 0, 'offset': 0}]


def extract_with_gemini(chunk_text: str, doc_id: str, paragraph_index: int, char_offset: int, custom_keys: Optional[List[str]] = None, api_key: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    Extract key-value pairs from a text chunk using Google Gemini LLM
    
    Args:
        chunk_text: Text chunk to process
        doc_id: Document identifier (e.g., D1, D2)
        paragraph_index: Starting paragraph index for this chunk
        char_offset: Character offset from document start
        custom_keys: Optional list of keys to prioritize
        api_key: Google Gemini API key
        
    Returns:
        list: Extracted records with key, value, raw_value, comments, provenance, confidence
    """
    # Build system prompt (per Final Plan template)
    system_prompt = """You are an extraction assistant. Given an input text chunk, return a JSON array of objects. Each object must have the fields:
- key (string): canonical name for the data point
- value (string): extracted value (preserve original words if possible)
- raw_value (string): exact substring from the input that supports this pair
- comments (string): DETAILED contextual commentary including: temporal context (as of what date/year), format notes (ISO dates, currency format), units of measurement, how this data point relates to other information, ambiguity notes, transformations applied, and analytical significance. Write 1-3 sentences of meaningful context.
- provenance (string): the exact sentence or phrase from the input text where this fact was found
- confidence (number): 0.0-1.0

IMPORTANT: The 'comments' field must be descriptive and informative, providing rich context about the data point. Never leave it empty.

Return only valid JSON. If no key/value pairs are present, return an empty array: []"""
    
    # Build user prompt (per Final Plan template)
    keys_hint = ""
    if custom_keys:
        keys_hint = f"\nPrioritize extracting these keys if present: {', '.join(custom_keys)}"
    
    user_prompt = f"""doc_id: {doc_id}
paragraph_index: {paragraph_index}
text: \"\"\"
{chunk_text}
\"\"\"

Rules:
1. Create separate objects for each distinct factual item (date, amount, name, address, phone, product, status, etc.). Do not omit any factual statements.
2. If a single sentence implies multiple keys (e.g., "John Doe, 42, lives at 12 Main St, pays $1200/month"), create separate objects for name, age, address, rent.
3. If a value is ambiguous, include both candidate values in 'value' separated by " | " and explain in comments.
4. Normalize dates to ISO (YYYY-MM-DD) in comments or raw_value only if you can be certain; otherwise keep original in value and note attempts in comments.
5. Do not hallucinate missing facts.{keys_hint}

CRITICAL: For EVERY extraction, provide detailed 'comments' explaining:
- Temporal context (e.g., "As of year 2024", "Current as of document date")
- Format details (e.g., "Formatted in ISO standard for easy parsing", "Currency in USD")
- Relationships (e.g., "Implies 15 years of experience since graduation", "Indicates senior-level position")
- Analytical significance (e.g., "Key demographic marker for analysis", "Critical for compliance tracking")

Example of good comments:
{{"key": "age", "value": "35", "raw_value": "35", "comments": "As of year 2024. This age serves as a key demographic marker for analytical purposes and indicates mid-career professional status.", "provenance": "...", "confidence": 0.95}}
{{"key": "salary", "value": "$85,000", "raw_value": "$85,000", "comments": "Annual compensation in USD. Formatted with currency symbol and comma separator. Indicates competitive mid-level compensation for the industry.", "provenance": "...", "confidence": 0.9}}

Return only the JSON array."""
    
    try:
        # Initialize Gemini client
        client = genai.Client(api_key=api_key)
        
        # Call Gemini API using google-genai library
        response = client.models.generate_content(
            model='gemini-2.0-flash-lite',
            contents=user_prompt,
            config=types.GenerateContentConfig(
                temperature=0.1,
                max_output_tokens=4096,
            ),
        )
        
        # Parse JSON response
        response_text = response.text if hasattr(response, 'text') else str(response)
        if response_text is None:
            response_text = ''
        response_text = response_text.strip()
        
        # Remove markdown code blocks if present
        if response_text.startswith('```'):
            parts = response_text.split('```')
            if len(parts) >= 2:
                response_text = parts[1]
                if response_text.startswith('json'):
                    response_text = response_text[4:]
            response_text = response_text.strip()
        
        extractions = json.loads(response_text)
        
        # Validate we got extractions
        if not extractions or not isinstance(extractions, list):
            # Log warning but don't fail - maybe text truly has no extractable data
            logger.warning(f"Gemini returned empty/invalid extraction for {doc_id}:para_{paragraph_index}. Text preview: {chunk_text[:200]}...")
            return []
        
        # Provenance is already in correct format from LLM
        # LLM returns: <doc_id>:<paragraph_index>:<char_start>-<char_end>
        for item in extractions:
            # If LLM didn't provide proper provenance, add default
            if not item.get('provenance') or ':' not in item.get('provenance', ''):
                item['provenance'] = f"{doc_id}:{paragraph_index}:0-{len(chunk_text)}"
        
        return extractions
        
    except json.JSONDecodeError as e:
        # Raise exception with helpful context
        error_text = response_text[:500] if 'response_text' in locals() and response_text else 'No response received'
        raise ValueError(f"Failed to parse JSON response from Gemini API. Response: {error_text}\nError: {str(e)}")
    except Exception as e:
        # Check for API key errors and provide clear message
        error_msg = str(e)
        if 'API key not valid' in error_msg or 'API_KEY_INVALID' in error_msg:
            raise ValueError("Invalid Gemini API key. Please check your API key and try again.\n\nTo get a valid API key:\n1. Go to https://aistudio.google.com/app/apikey\n2. Create or copy your API key\n3. Paste it in the sidebar")
        elif 'INVALID_ARGUMENT' in error_msg:
            raise ValueError(f"Invalid API request: {error_msg}")
        else:
            raise Exception(f"Gemini API call failed: {error_msg}")


def convert_to_three_columns(extractions: List[Dict[str, Any]]) -> List[Dict[str, str]]:
    """
    Convert Gemini extraction format to final 3-column format: key, value, comments
    
    Args:
        extractions: List of extraction dicts from Gemini (with raw_value, provenance, confidence)
        
    Returns:
        list: List of dicts with only 3 keys: key, value, comments
    """
    final_rows = []
    
    for item in extractions:
        # Get core fields
        key = item.get('key', 'unknown_key')
        value = item.get('value', '')
        
        # Use only the LLM's comments field for the comments column
        # This contains context, units, ambiguity notes, and transformations
        comments = item.get('comments', '')
        
        final_rows.append({
            'key': key,
            'value': value,
            'comments': comments
        })
    
    return final_rows
