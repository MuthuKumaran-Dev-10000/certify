import pandas as pd

# Create sample data for 10 rows
data = {
    "Name": ["Muthu", "Ajay", "Mathesh", "Santhosh", "Hari", "Arun", "Kumar", "Divya", "Sneha", "Gokul"],
    "Event": ["Algogeeks", "CodeRush", "HackHub", "TechSprint", "BugHunt", "InnoFest", "CodeMania", "DataDive", "AIStorm", "TechTrek"],
    "Prize": ["1st Prize", "2nd Prize", "3rd Prize", "1st Prize", "Participation", "2nd Prize", "1st Prize", "Participation", "3rd Prize", "2nd Prize"],
    "Date": ["18.08.2025"] * 10
}

# Convert to DataFrame
df = pd.DataFrame(data)

# Save to Excel file
file_path = "sample_certify_data.xlsx"
df.to_excel(file_path, index=False)

file_path
