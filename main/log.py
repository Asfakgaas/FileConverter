import pandas as pd
input_file = pd.read_csv(file_path)
output_file = pd.ExcelWriter()
input_file.to_excel(output_file, index=False)
output_file.save()