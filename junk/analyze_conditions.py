import pandas as pd
import re
import json

try:
    df = pd.read_excel('noms_38.xlsx')
    names = df['Название'].dropna().unique().tolist()
    
    # Common condition patterns: c1, c2, c3, p9, rb, co, etc.
    # Usually they are at the end of the string after a space.
    
    patterns = {}
    for name in names:
        parts = name.split()
        if len(parts) > 1:
            suffix = parts[-1].lower()
            # Heuristic: if suffix starts with c, p, or is a small alphanumeric
            if re.match(r'^[cp]\d+$|^rb$|^co$|^\d+/\d+$', suffix):
                patterns[suffix] = patterns.get(suffix, 0) + 1
    
    with open('condition_analysis.json', 'w', encoding='utf-8') as f:
        json.dump(patterns, f, ensure_ascii=False, indent=2)
    
    print("Analysis done. Results in condition_analysis.json")

except Exception as e:
    print(f"Error: {e}")
