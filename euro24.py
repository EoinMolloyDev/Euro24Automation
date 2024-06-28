import pandas as pd
import numpy as np

# Define the path to the Excel file
file_path = 'euro24.xlsx'

# Load the Excel file and get all sheet names
excel_file = pd.ExcelFile(file_path)
sheet_names = excel_file.sheet_names

# Exclude specific sheet names from the list of sheets to process
exclude_sheets = ['Fund', 'Standings', 'Results Sheet']
other_sheets = [sheet for sheet in sheet_names if sheet not in exclude_sheets]

# Load the master sheet
master_df = pd.read_excel(file_path, sheet_name='Results Sheet', header=1)

# Remove non-numeric rows from the score columns and fill NaN values with 0
master_df = master_df[pd.to_numeric(master_df.iloc[:, 2], errors='coerce').notnull()]
master_df = master_df[pd.to_numeric(master_df.iloc[:, 3], errors='coerce').notnull()]
master_df.iloc[:, 2] = master_df.iloc[:, 2].fillna(0).astype(int)
master_df.iloc[:, 3] = master_df.iloc[:, 3].fillna(0).astype(int)
master_df.iloc[:, 5] = master_df.iloc[:, 5].fillna(0)

# Initialize a dictionary to store points
points = {sheet: 0 for sheet in other_sheets}

# Function to compare scores and assign points
def compare_scores(master_row, other_row, sheet):
    points_awarded = 0
    
    master_team1_score = master_row.iloc[2]
    master_team2_score = master_row.iloc[3]
    other_team1_score = other_row.iloc[2]
    other_team2_score = other_row.iloc[3]
    
    # Determine the results as -1, 0, 1 for loss, draw, win
    if master_team1_score > master_team2_score:
        master_result = 1
    elif master_team1_score < master_team2_score:
        master_result = -1
    else:
        master_result = 0

    if other_team1_score > other_team2_score:
        other_result = 1
    elif other_team1_score < other_team2_score:
        other_result = -1
    else:
        other_result = 0
    
    if master_result == other_result:
        points_awarded += 2
        #print(f"{sheet}: Correct result for {master_row[1]} vs {master_row[4]} on {master_row[0]} - {master_team1_score}:{master_team2_score}")
    
    # Check if the exact score is correct
    if master_team1_score == other_team1_score and master_team2_score == other_team2_score:
        points_awarded += 1
        #print(f"{sheet}: Exact score for {master_row[1]} vs {master_row[4]} on {master_row[0]} - {master_team1_score}:{master_team2_score}")
    
    return points_awarded

# Function to find the closest time to the first goal
def find_closest_first_goal(master_time, all_times):
    # Ensure that all times are valid integers
    valid_times = [(sheet, int(time)) for sheet, time in all_times if pd.notnull(time) and str(time).isdigit()]
    if not valid_times:  # Check if the list is empty after filtering
        return []
    closest_time = min(valid_times, key=lambda x: abs(x[1] - master_time))[1]
    closest_sheets = [sheet for sheet, time in valid_times if time == closest_time]
    return closest_sheets

# Collect all first goal times across sheets
all_first_goal_times = []

# Process each row in the master sheet
for index, master_row in master_df.head(38).iterrows():
    date = master_row[0]
    team1 = master_row[1]
    team1Score = master_row[2]
    team2 = master_row[4]
    team2Score = master_row[3]
    master_first_goal_time = master_row[5]
    #print(f"Processing row {index} for {date}, {team1} {team1Score} - {team2Score} {team2}")
    
    # Collect first goal times across sheets for the current row
    all_first_goal_times = []
    
    # Process each other sheet
    for sheet in other_sheets:
        other_df = pd.read_excel(file_path, sheet_name=sheet, header=1)
        
        # Remove non-numeric rows from the score columns and fill NaN values with 0
        other_df = other_df[pd.to_numeric(other_df.iloc[:, 2], errors='coerce').notnull()]
        other_df = other_df[pd.to_numeric(other_df.iloc[:, 3], errors='coerce').notnull()]
        other_df.iloc[:, 2] = other_df.iloc[:, 2].fillna(0).astype(int)
        other_df.iloc[:, 3] = other_df.iloc[:, 3].fillna(0).astype(int)
        other_df.iloc[:, 5] = other_df.iloc[:, 5].fillna(0)
        
        # Find the corresponding row in the other sheet
        other_row = other_df[(other_df.iloc[:, 0] == date) & (other_df.iloc[:, 1] == team1) & (other_df.iloc[:, 4] == team2)]
        
        if not other_row.empty:
            points[sheet] += compare_scores(master_row, other_row.iloc[0], sheet)
            # Collect first goal times for closest time calculation
            all_first_goal_times.append((sheet, other_row.iloc[0, 5]))
    
    # Determine the closest time for the first goal for the current row
    if all_first_goal_times:
        closest_sheets = find_closest_first_goal(master_first_goal_time, all_first_goal_times)
        for sheet in closest_sheets:
            points[sheet] += 2
            print(f"{sheet}: Closest time for first goal for {team1} vs {team2} on {date} - {master_first_goal_time}")

# Display the points sorted by most points
sorted_points = sorted(points.items(), key=lambda item: item[1], reverse=True)
for sheet, point in sorted_points:
    print(f"{sheet}: {point} points")
