import pandas as pd
import matplotlib.pyplot as plt

# Excel file
file_path = 'data/Daniel.Noa-HamiltonCC-05.xlsx'
data = pd.read_excel(file_path)

# This cleans the data by removing any rows that aren't needed (like headers and empty rows)
data_cleaned = data.dropna().reset_index(drop=True)

# Extract the relevant columns: Class Name and Total Enrollment
class_data = data_cleaned.iloc[1:, [0, -1]]  # Assuming first column is Class Name, last column is Total
class_data.columns = ['Class Name', 'Total Enrollment']

# Convert the 'Total Enrollment' column to numeric (in case it's not)
class_data['Total Enrollment'] = pd.to_numeric(class_data['Total Enrollment'])

# Display the cleaned data
print("Cleaned Data:")
print(class_data)

# Visualize the total enrollment per class using a bar chart
class_data.plot(kind='bar', x='Class Name', y='Total Enrollment', title='Total Enrollment per Class', ylabel='Total Enrollment', xlabel='Class Name')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
plt.show()

# Save the plot as an image file
plt.savefig('total_enrollment_per_class.png')
