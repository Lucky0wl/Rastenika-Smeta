import pandas as pd
import re
import json

try:
    # Use the backup or the current file (current has part of it moved already)
    # Actually, current 'noms_38.xlsx' has some names cleaned. 
    # I should check if '100-120' is still in 'Название'.
    df = pd.read_excel('noms_38.xlsx')
    names = df['Название'].dropna().unique().tolist()
    
    size_patterns = {}
    for name in names:
        # Look for patterns like 100-120, 100/120, 120+, 80-100
        # at the end of the string
        parts = name.split()
        if not parts: continue
        
        last_part = parts[-1]
        if re.search(r'\d+[-/]\d+|\d+\+|\d+-\d+-\d+', last_part):
            size_patterns[last_part] = size_patterns.get(last_part, 0) + 1
            
    with open('size_analysis.json', 'w', encoding='utf-8') as f:
        json.dump(size_patterns, f, ensure_ascii=False, indent=2)
    
    print(f"Found {len(size_patterns)} unique size patterns. Results in size_analysis.json")

except Exception as e:
    print(f"Error: {e}")
