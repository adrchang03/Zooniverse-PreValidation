import pandas as pd
import json
from tkinter import Tk, simpledialog, messagebox
from tkinter.filedialog import askopenfilename, asksaveasfilename

# Use tkinter to open a file explorer popup for selecting the input file
def get_file_via_explorer():
    Tk().withdraw()  # We don't want a full GUI, so we keep the root window from appearing
    file_path = askopenfilename(title="Select the Excel file", filetypes=[("Excel files", "*.xlsx *.xls")])
    return file_path

# Use tkinter to open a file explorer popup for saving the output file
def get_save_file_location():
    Tk().withdraw()  # We don't want a full GUI, so we keep the root window from appearing
    save_path = asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], title="Save the extracted file")
    return save_path

# Function to extract data from 'subject_data' (column M)
def extract_filename(cell):
    try:
        data = json.loads(cell)
        # Extract the first key and then go deeper to get the filename
        first_key = list(data.keys())[0]
        return data[first_key]["Filename"]
    except (json.JSONDecodeError, KeyError):
        return None

# Function to extract data from 'annotations' (column L)
def extract_annotations(cell):
    species_data = {
        "species_1": "NONE", "how_many_1": "NONE",
        "species_2": "NONE", "how_many_2": "NONE",
        "time_of_day": "NODATA", "temperature": "NODATA",
        "month": "NODATA", "habitat": "NODATA"
    }

    try:
        data = json.loads(cell)
        # Extract species and related info
        if "value" in data[0]:
            for i, item in enumerate(data[0]["value"]):
                species_key = f"species_{i+1}"
                how_many_key = f"how_many_{i+1}"
                
                species_data[species_key] = item.get("choice", "NONE")
                species_data[how_many_key] = item.get("answers", {}).get("HOWMANY", "NONE")
                
                # Other attributes (only take from the first species)
                if i == 0:
                    species_data["time_of_day"] = item.get("answers", {}).get("TIMEOFDAY", "NODATA")
                    species_data["temperature"] = item.get("answers", {}).get("TEMPERATURE", "NODATA")
                    species_data["month"] = item.get("answers", {}).get("MONTHOFTHEYEAR", "NODATA")
                    species_data["habitat"] = item.get("answers", {}).get("HABITAT", "")
    except (json.JSONDecodeError, KeyError):
        pass

    return species_data

# Prompt user for starting row number
def ask_for_start_row():
    Tk().withdraw()  # Keep the root window hidden
    start_row = simpledialog.askinteger("Input", "Enter the starting row number:", minvalue=1)
    return start_row

# Prompt user for column F text
def ask_for_workflow_name():
    Tk().withdraw()
    workflow_name = simpledialog.askstring("Input", "Enter the exact text to match in column F (workflow_name):")
    return workflow_name

# Load the selected Excel file - MAIN LOGIC
input_file = get_file_via_explorer()
if input_file:
    df = pd.read_excel(input_file)
    print(f"Loaded file: {input_file}")
    
    # Ask the user to input the starting row number
    start_row = ask_for_start_row()
    if start_row is None:
        messagebox.showerror("Error", "No starting row specified. Exiting the program.")
        exit()


    # Ask the user for workflow name
    workflow_name = ask_for_workflow_name()
    if not workflow_name:
        messagebox.showerror("Error", "No workflow name specified. Exiting the program.")
        exit()


    # Create new DataFrame to store extracted data
    extracted_data = {
        "FILENAME": [],
        "SPECIES 1": [],
        "HOW MANY OF SPECIES 1": [],
        "SPECIES 2": [],
        "HOW MANY OF SPECIES 2": [],
        "TIME OF DAY": [],
        "TEMPERATURE": [],
        "MONTH": [],
        "HABITAT": []
    }

    # Adjust the DataFrame to start from the specified row (start_row-1 because it's 0-indexed)
    df = df.iloc[start_row-1:]

    # Iterate through each row to extract relevant information
    for index, row in df.iterrows():
        # Only process the row if column F has the exact text provided from pop up
        if row.get('workflow_name') != workflow_name:
            continue  # Skip this row if the condition is not met
        
        # Extract filename from 'subject_data' column (column M)
        filename = extract_filename(row['subject_data'])
        
        # Extract species and other info from 'annotations' column (column L)
        annotations = extract_annotations(row['annotations'])
        
        # Append data to the extracted_data dictionary
        extracted_data["FILENAME"].append(filename)
        extracted_data["SPECIES 1"].append(annotations["species_1"])
        extracted_data["HOW MANY OF SPECIES 1"].append(annotations["how_many_1"])
        extracted_data["SPECIES 2"].append(annotations["species_2"])
        extracted_data["HOW MANY OF SPECIES 2"].append(annotations["how_many_2"])
        extracted_data["TIME OF DAY"].append(annotations["time_of_day"])
        extracted_data["TEMPERATURE"].append(annotations["temperature"])
        extracted_data["MONTH"].append(annotations["month"])
        extracted_data["HABITAT"].append(annotations["habitat"])

    # Convert the extracted data to a new DataFrame
    extracted_df = pd.DataFrame(extracted_data)

    # Save the extracted data into a new Excel file
    save_file = get_save_file_location()  # Get location for saving the file
    if save_file:
        extracted_df.to_excel(save_file, index=False, startcol=11)
        print(f"Data extracted and saved to {save_file}")
    else:
        print("No file selected for saving.")
else:
    print("No file selected for input.")
