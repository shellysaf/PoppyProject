import pandas as pd
from scipy.stats import ttest_rel
from itertools import combinations

# Load the data from the Excel file
file_path = f'C:\\Users\\SASH723\\Documents\\לימודים\\פרוייקט גמר\\ניתוחים נוספים\\successes analysis.xlsx'
df = pd.read_excel(file_path, sheet_name='DB')

# Aggregate data: mean success rate per participant per angle across all exercises
mean_success = df.groupby(['Participant', 'Angle'])['Num of success'].mean().reset_index()

# Get the unique angles
angles = mean_success['Angle'].unique()

# Step 1: Define the list to store the results
results = []

# Step 2: Iterate over all combinations of angle pairs
for angle1, angle2 in combinations(angles, 2):
    # Step 2a: Extract the data for the two angles
    data1 = mean_success[mean_success['Angle'] == angle1]['Num of success']
    data2 = mean_success[mean_success['Angle'] == angle2]['Num of success']

    # Step 2b: Perform paired t-test
    t_stat, p_value = ttest_rel(data1, data2)

    # Step 2c: Apply Bonferroni correction: multiply p-value by the number of comparisons
    p_value_corrected = p_value * len(list(combinations(angles, 2)))

    # Step 2d: Append results to the list
    results.append((angle1, angle2, t_stat, p_value, p_value_corrected))

# Step 3: Convert results to a DataFrame for easier viewing
results_df = pd.DataFrame(results, columns=['Angle 1', 'Angle 2', 'T-Statistic', 'P-Value', 'P-Value Corrected'])

# Step 4: Display the results
print(results_df)
