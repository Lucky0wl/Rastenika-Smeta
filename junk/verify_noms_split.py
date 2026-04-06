import pandas as pd

try:
    # Read the main sheet
    df = pd.read_excel('noms_38.xlsx')
    print("Main Sheet Columns:", df.columns.tolist())
    
    # Read the conditions sheet
    df_conds = pd.read_excel('noms_38.xlsx', sheet_name='Кондиции')
    print("Conditions Sheet Count:", len(df_conds))
    print("Unique Conditions:", df_conds['Кондиция'].tolist()[:10])

    # Check first 5 rows of main sheet
    print("\nSample Data (First 5):")
    sample = df[['Название', 'Кондиция']].head(5)
    print(sample)
    
    # Check if we have some populated conditions
    populated_conds = df[df['Кондиция'] != ""]
    print(f"\nRows with populated conditions: {len(populated_conds)}")
    
    if len(populated_conds) > 0:
        print("Verification SUCCESS")
    else:
        print("Verification FAILURE: No conditions extracted")

except Exception as e:
    print(f"Error during verification: {e}")
