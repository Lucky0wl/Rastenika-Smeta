import pandas as pd

try:
    df = pd.read_excel('noms_38.xlsx')
    df_conds = pd.read_excel('noms_38.xlsx', sheet_name='Кондиции')
    
    print(f"Total rows: {len(df)}")
    print(f"Unique conditions count: {len(df_conds)}")
    
    # Check for sizes in conditions
    conds = df['Кондиция'].dropna().astype(str).unique().tolist()
    sizes = [c for c in conds if '-' in c or '+' in c]
    print(f"\nConditions with sizes identified: {len(sizes)}")
    print("Sample sizes:", sizes[:10])
    
    # Check sample rows
    print("\nSample Data (First 10):")
    print(df[['Название', 'Кондиция']].head(10))
    
    # Find some rows that likely had multiple parameters
    multi = df[df['Кондиция'].str.contains(' ', na=False)]
    if not multi.empty:
        print("\nRows with multiple condition parameters (e.g. c2 100-120):")
        print(multi[['Название', 'Кондиция']].head(5))
    
    print("\nVerification SUCCESS")
except Exception as e:
    print(f"Error: {e}")
