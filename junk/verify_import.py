import pandas as pd
import os

def test_import():
    filepath = 'noms_38.xlsx'
    if not os.path.exists(filepath):
        print(f"Error: {filepath} not found")
        return

    try:
        df = pd.read_excel(filepath)
        col_map = {
            'Наименование': ['Наименование', 'Растение', 'Название', 'Name', 'Прайс', 'Товар'],
            'Кондиция': ['Кондиция', 'Параметры', 'Размер', 'Характеристики', 'Описание', 'Condition', 'Сорт'],
            'Цена': ['Цена', 'Стоимость', 'Price', 'Цена базового материала', 'Опт']
        }
        
        found_cols = {}
        for target, aliases in col_map.items():
            for col in df.columns:
                if any(alias.lower() in str(col).lower() for alias in aliases):
                    found_cols[target] = col
                    break
        
        print(f"Found columns: {found_cols}")
        
        if 'Наименование' not in found_cols:
            print("Error: Name column not found")
            return

        plants = []
        for _, row in df.iterrows():
            name = str(row[found_cols['Наименование']])
            if name == 'nan' or not name.strip(): continue
            plants.append(name)
        
        print(f"Successfully parsed {len(plants)} plants. First 5: {plants[:5]}")
    except Exception as e:
        print(f"Error during import: {e}")

if __name__ == '__main__':
    test_import()
