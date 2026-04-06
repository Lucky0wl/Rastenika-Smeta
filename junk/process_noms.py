import pandas as pd
import re
import os

def split_condition(name):
    if not isinstance(name, str) or not name.strip():
        return name, ""
    
    # Common patterns: 
    # c2, p9, rb, c15, c3, c5, c10, c20, rb, c25, c12, c35
    # Also handle things like 100/120 (height) or RB (Root Ball)
    # We look for the last part of the name if it looks like a container/size
    
    parts = name.strip().split()
    if len(parts) < 2:
        return name, ""
    
    last_word = parts[-1]
    
    # Pattern 1: [letter][number] e.g. c2, p9, c15, pa80
    # Pattern 2: rb, co (case insensitive)
    # Pattern 3: number/number e.g. 100/120
    
    pattern1 = re.compile(r'^[a-zA-Z]+\d+$', re.IGNORECASE)
    pattern2 = re.compile(r'^(rb|co|br|mult)$', re.IGNORECASE)
    pattern3 = re.compile(r'^\d+/\d+$')
    
    if pattern1.match(last_word) or pattern2.match(last_word) or pattern3.match(last_word):
        new_name = " ".join(parts[:-1]).strip()
        condition = last_word
        return new_name, condition
    
    return name, ""

def process_file(input_file, output_file):
    print(f"Reading {input_file}...")
    # Read all sheets (if any) - but we mostly care about the first one
    all_sheets = pd.read_excel(input_file, sheet_name=None)
    
    # Get the first sheet name
    sheet_name = list(all_sheets.keys())[0]
    df = all_sheets[sheet_name]
    
    if 'Название' not in df.columns:
        print("Error: 'Название' column not found.")
        return

    print("Splitting conditions...")
    results = df['Название'].apply(split_condition)
    df['Название'] = [r[0] for r in results]
    df['Кондиция'] = [r[1] for r in results]
    
    # Reorder columns: put 'Кондиция' after 'Название'
    cols = list(df.columns)
    name_idx = cols.index('Название')
    cols.insert(name_idx + 1, cols.pop(cols.index('Кондиция')))
    df = df[cols]
    
    # Get unique conditions for the separate table
    unique_conditions = df['Кондиция'].unique()
    unique_conditions = [c for c in unique_conditions if c] # remove empty
    cond_df = pd.DataFrame({'Кондиция': unique_conditions})
    
    print(f"Saving to {output_file}...")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        cond_df.to_excel(writer, index=False, sheet_name='Кондиции')
    
    print("Success!")

if __name__ == "__main__":
    input_path = 'noms_38.xlsx'
    output_path = 'noms_38.xlsx' # Overwrite as requested (separate means structure change)
    # Safer to backup first
    if os.path.exists(input_path):
        os.rename(input_path, 'noms_38_backup.xlsx')
        process_file('noms_38_backup.xlsx', output_path)
    else:
        print(f"File {input_path} not found.")
