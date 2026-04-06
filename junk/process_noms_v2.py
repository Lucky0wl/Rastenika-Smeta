import pandas as pd
import re
import os

def extract_conditions(name):
    if not isinstance(name, str) or not name.strip():
        return name, ""
    
    # Patterns to match:
    # 1. Container: c2, p9, rb, c15, pa80, co2, c25, etc.
    pattern_container = re.compile(r'^[a-zA-Z]+\d+$|^rb$|^co$|^br$|^mult$|^rb/c\d+$', re.IGNORECASE)
    # 2. Size: 100-120, 100/120, 40+, 180-200см, 3-4лет
    pattern_size = re.compile(r'^\d+[-/]\d+[а-яА-Я]*$|^\d+\+$|^\d+-\d+-\d+$|^\d+[а-яА-Я]+$', re.IGNORECASE)
    
    parts = name.strip().split()
    extracted = []
    
    # Iterate from the end and pull out anything that looks like a condition
    while parts:
        last_word = parts[-1]
        # Clean the word for matching (remove trailing dots/commas if any)
        clean_word = last_word.strip('.,')
        
        if pattern_container.match(clean_word) or pattern_size.match(clean_word):
            extracted.insert(0, last_word) # Keep original formatting
            parts.pop()
        else:
            break
            
    new_name = " ".join(parts).strip()
    condition = " ".join(extracted).strip()
    
    return new_name, condition

def process_file(input_file, output_file):
    print(f"Reading {input_file}...")
    all_sheets = pd.read_excel(input_file, sheet_name=None)
    sheet_name = list(all_sheets.keys())[0]
    df = all_sheets[sheet_name]
    
    if 'Название' not in df.columns:
        print("Error: 'Название' column not found.")
        return

    print("Extracting conditions and sizes...")
    results = df['Название'].apply(extract_conditions)
    df['Название'] = [r[0] for r in results]
    new_conds = [r[1] for r in results]
    
    # If 'Кондиция' already exists (e.g. from previous run), merge them
    if 'Кондиция' in df.columns:
        df['Кондиция'] = df['Кондиция'].fillna('').astype(str)
        # Merge: existing + space + new
        df['Кондиция'] = [ (str(e).strip() + " " + str(n).strip()).strip() for e, n in zip(df['Кондиция'], new_conds)]
    else:
        df['Кондиция'] = new_conds
        # Reorder: put 'Кондиция' after 'Название'
        cols = list(df.columns)
        name_idx = cols.index('Название')
        cols.insert(name_idx + 1, cols.pop(cols.index('Кондиция')))
        df = df[cols]
    
    # Generate unique conditions list for the 'Кондиции' sheet
    unique_conditions = df['Кондиция'].unique()
    unique_conditions = [c for c in unique_conditions if c and c.lower() != 'nan']
    cond_df = pd.DataFrame({'Кондиция': sorted(unique_conditions)})
    
    print(f"Saving to {output_file}...")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        cond_df.to_excel(writer, index=False, sheet_name='Кондиции')
    
    print("Success!")

if __name__ == "__main__":
    # Use backup if available to start fresh with improved logic
    input_path = 'noms_38_backup.xlsx'
    if not os.path.exists(input_path):
        input_path = 'noms_38.xlsx'
        
    output_path = 'noms_38.xlsx'
    process_file(input_path, output_path)
