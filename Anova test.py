import pandas as pd
from statsmodels.stats.anova import AnovaRM

# Step 1: Load the data from the Excel file
file_path = f'C:\\Users\\SASH723\\Documents\\לימודים\\פרוייקט גמר\\ניתוחים נוספים\\successes analysis.xlsx'
df = pd.read_excel(file_path, sheet_name='DB')

# Step 2: Aggregate data: mean success rate per participant per angle across all exercises
mean_success = df.groupby(['Participant', 'Angle'])['Num of success'].mean().reset_index()

# Step 3: Perform Repeated Measures ANOVA
aovrm = AnovaRM(mean_success, 'Num of success', 'Participant', within=['Angle'])
anova_results = aovrm.fit()

# Step 4: Display the ANOVA results
print("Repeated Measures ANOVA Results:")
print(anova_results.summary())
