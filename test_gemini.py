"""
Test script to validate Gemini LLM extraction
"""

import sys
from utils.parser import parse_key_value_pairs

# Read test data
with open('ref/Data Input.txt', 'r', encoding='utf-8') as f:
    test_text = f.read()

# Get API key from command line
if len(sys.argv) < 2:
    print("Usage: python test_gemini.py <API_KEY>")
    sys.exit(1)

api_key = sys.argv[1]

print("Testing Gemini extraction...")
print(f"Input text length: {len(test_text)} characters")
print("-" * 80)

try:
    # Extract data
    results = parse_key_value_pairs(test_text, api_key, custom_keys=None)
    
    print(f"\n✅ Extraction successful! Found {len(results)} records.\n")
    
    # Display first 10 results
    for i, record in enumerate(results[:10], 1):
        print(f"\n{i}. Key: {record['key']}")
        print(f"   Value: {record['value']}")
        print(f"   Comments: {record['comments'][:100]}..." if len(record.get('comments', '')) > 100 else f"   Comments: {record.get('comments', '')}")
    
    if len(results) > 10:
        print(f"\n... and {len(results) - 10} more records")
    
    print("\n" + "=" * 80)
    print("Test completed successfully!")
    
except Exception as e:
    print(f"\n❌ Error during extraction: {str(e)}")
    import traceback
    traceback.print_exc()
