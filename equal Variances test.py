import pandas as pd
from scipy.stats import levene

# Load your data from the Excel file
file_path = f'C:\\Users\\SASH723\\Documents\\לימודים\\פרוייקט גמר\\ניתוחים נוספים\\successes analysis.xlsx'

# Load the data from the 'DB' sheet
df = pd.read_excel(file_path, sheet_name='DB')

# Aggregate data: mean success rate per participant per angle across all exercises
mean_success = df.groupby(['Participant', 'Angle'])['Num of success'].mean().reset_index()

# Get the unique angles
angles = mean_success['Angle'].unique()

# Prepare the data for Levene's Test by grouping the success rates for each angle
grouped_data = [mean_success[mean_success['Angle'] == angle]['Num of success'] for angle in angles]

# Perform Levene's Test for Equal Variances across angles
levene_stat, levene_p_value = levene(*grouped_data)

# Display the results of Levene's Test
print("Levene's Test for Equal Variances:")
print(f"Levene Statistic: {levene_stat:.4f}")
print(f"p-Value: {levene_p_value:.4f}")

# Interpretation of the results
if levene_p_value > 0.05:
    print("Equal variances can be assumed.")
else:
    print("Equal variances cannot be assumed.")
