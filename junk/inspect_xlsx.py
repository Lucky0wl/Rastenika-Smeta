import pandas as pd
import json

try:
    df = pd.read_excel('noms_38.xlsx')
    info = {
        "columns": df.columns.tolist(),
        "head": df.head(10).to_dict(orient='records')
    }
    with open('inspect_results.json', 'w', encoding='utf-8') as f:
        json.dump(info, f, ensure_ascii=False, indent=2)
    print("Done. Results saved to inspect_results.json")
except Exception as e:
    print(f"Error: {e}")
