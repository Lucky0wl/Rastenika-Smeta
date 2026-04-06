import openpyxl
import os

filename = 'Шаблон.xlsx'
if os.path.exists(filename):
    try:
        wb = openpyxl.load_workbook(filename)
        ws = wb.active
        
        def get_all_styles(cell_coord):
            cell = ws[cell_coord]
            return {
                'val': cell.value,
                'font': {
                    'name': cell.font.name,
                    'sz': cell.font.size,
                    'b': cell.font.bold,
                    'i': cell.font.italic,
                    'color': cell.font.color.rgb if cell.font.color and hasattr(cell.font.color, 'rgb') else '000000'
                },
                'fill': {
                    'fgColor': cell.fill.start_color.rgb if cell.fill and hasattr(cell.fill.start_color, 'rgb') else 'FFFFFF'
                },
                'align': {
                    'h': cell.alignment.horizontal,
                    'v': cell.alignment.vertical
                },
                'border': {
                    'top': cell.border.top.style,
                    'bottom': cell.border.bottom.style,
                    'left': cell.border.left.style,
                    'right': cell.border.right.style
                }
            }

        # Analysis of key cells
        # Title (A2 likely)
        print("--- A2 Style ---")
        print(get_all_styles('A2'))
        
        # Header (A6 likely)
        print("\n--- A6 Style ---")
        print(get_all_styles('A6'))
        
        # Global Row Heights
        print("\n--- Row Heights (1-10) ---")
        for i in range(1, 11):
            print(f"Row {i}: {ws.row_dimensions[i].height}")
            
        # Merged cells
        print("\n--- Merged Cells ---")
        print(ws.merged_cells.ranges)
        
    except Exception as e:
        print(f"Error: {e}")
else:
    print(f"File {filename} not found.")
