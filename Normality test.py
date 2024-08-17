import pandas as pd
from scipy.stats import shapiro

# Load your data from the Excel fileode
file_path = f'C:\\Users\\SASH723\\Documents\\לימודים\\פרוייקט גמר\\ניתוחים נוספים\\successes analysis.xlsx'

# Load the data from the 'DB' sheet
df = pd.read_excel(file_path, sheet_name='DB')

# Aggregate data: mean success rate per participant per angle across all exercises
mean_success = df.groupby(['Participant', 'Angle'])['Num of success'].mean().reset_index()

print(mean_success)

# Perform the Shapiro-Wilk test for normality on the success rates for each angle
angles = mean_success['Angle'].unique()
normality_tests = {angle: shapiro(mean_success[mean_success['Angle'] == angle]['Num of success']) for angle in angles}

# Display the results
normality_tests_results = {angle: {'W-Statistic': result.statistic, 'p-Value': result.pvalue} for angle, result in normality_tests.items()}

# Print the results
for angle, results in normality_tests_results.items():
    print(f"Angle: {angle}°")
    print(f"  W-Statistic: {results['W-Statistic']:.4f}")
    print(f"  p-Value: {results['p-Value']:.4f}")
    print()
