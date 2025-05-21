import pandas as pd

input_file_path = r"H:\Reference\SerialNumber\output_with_start_dates.xlsx"
df = pd.read_excel(input_file_path)

# Convert the 'Start Date' column to datetime format
df['Start Date'] = pd.to_datetime(df['Start Date'], errors='coerce')

# Format dates as MM/DD/YYYY
df['Start Date'] = df['Start Date'].dt.strftime('%m/%d/%Y')

# Save the updated DataFrame to a new Excel file
output_file_path = r"H:\Reference\SerialNumber\output_converted_start_dates.xlsx"
df.to_excel(output_file_path, index=False)

print(f"Process completed. Results saved to {output_file_path}.")