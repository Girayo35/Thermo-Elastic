import os
import pandas as pd

folder_path = 'Node_Information'  # Your folder path
output_file = 'Collected_Integers_By_File.xlsx'

column_list = []

for file in os.listdir(folder_path):
    if file.endswith('.xlsx'):
        file_path = os.path.join(folder_path, file)
        file_base = os.path.splitext(file)[0]
        collected_values = []

        try:
            xls = pd.ExcelFile(file_path)
            for sheet in xls.sheet_names:
                df = xls.parse(sheet)

                # Ensure column B exists and starts from at least row 13
                if df.shape[0] >= 13 and df.shape[1] >= 2:
                    col_data = df.iloc[12:, 1]  # B13 downward (no max row)

                    # Convert to numeric and keep only valid integers
                    col_data_numeric = pd.to_numeric(col_data, errors='coerce')
                    int_values = col_data_numeric[(col_data_numeric.notna()) & (col_data_numeric % 1 == 0)]
                    int_values = int_values.astype(int)

                    collected_values.extend(int_values.tolist())

        except Exception as e:
            print(f"❌ Error in '{file}': {e}")
            continue

        if collected_values:
            col_df = pd.DataFrame({file_base: collected_values})
            column_list.append(col_df)
        else:
            print(f"⚠️ No integer values found in {file}")

# === Combine all columns into one DataFrame
if column_list:
    final_df = pd.concat(column_list, axis=1)
    final_df.to_excel(output_file, index=False)
    print(f"\n✅ Final Excel saved (all rows): {output_file}")
else:
    print("\n❌ No data collected from any file.")
