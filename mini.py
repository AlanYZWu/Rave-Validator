import openpyxl as xl


def extract_oid(string):
    copy = False  # Indicates whether we're reading something in []
    oid_temp = ""  # OID we're copying
    oid_list_temp = list()  # List of OIDs in description
    for i in range(len(string)):  # Iterates over string
        if copy and string[i] != "]":  # Copies characters onto OID string
            oid_temp += string[i]
        elif string[i] == "[":  # Indicates we're iterating over an OID
            copy = True
        elif string[i] == "]":  # Indicates that we're done iterating over an OID
            copy = False
            oid_list_temp.append(oid_temp)  # Inserts OID into list
            oid_temp = ""

    return oid_list_temp


def number_to_letter(num):
    return chr(num + 64)


# Load workbooks
sds_book = xl.load_workbook(filename="US_Global_Library_GL_V1.00.xlsx")
dynamics_book = xl.load_workbook(filename="Safety Matrix CRF Design Supplement_V0.3.xlsm")

# Get Dynamics sheet from Dynamics and CheckSteps from Rave SDS
dynamics_sheet = dynamics_book["Dynamics"]
check_steps_sheet = sds_book["CheckSteps"]

# Get column D (Edit Check) from Dynamics and column A (CheckName) from CheckSteps and subset to relevant rows
edit_checks_column = dynamics_sheet["D"]
edit_checks_column = edit_checks_column[2:]

check_names_column = check_steps_sheet["A"]
check_names_column = check_names_column[1:]

# Create output file
output = open("output.txt", "w+")
output.write("Edit Checks Not Found:\n")

# Search for corresponding edit check programmed in CheckSteps
for edit_check_cell in edit_checks_column:  # Iterates over each edit check
    if edit_check_cell.value is None:  # Stops when blank cell is found
        break
    edit_check_split = edit_check_cell.value.split("\n")  # Separates each dynamic in the cell
    for edit_check in edit_check_split:  # Iterates over each dynamic found in Edit Check cell
        found = False  # Tracks if edit_check is found
        if "_" not in edit_check:  # Skips if edit check doesn't have _ meaning it's not a dynamic
            continue
        for check_name_cell in check_names_column:  # Iterates over each check in CheckStep
            if check_name_cell.value == edit_check:
                found = True
                break
        if not found:  # Write to output file which edit checks weren't found
            output.write(f"{edit_check} in {number_to_letter(edit_check_cell.column)}"
                         f"{edit_check_cell.row} was not found\n")

# Check each description for correct folder and formOID
output.write("\nIncorrect Dynamics Descriptions:\n")
for edit_check_cell in edit_checks_column:  # Iterates through all edit checks
    if edit_check_cell.value is None:  # stops when blank cell is found
        break
    edit_check_split = edit_check_cell.value.split("\n")  # Separate each check in the cell
    for edit_check in edit_check_split:
        if "_" not in edit_check:  # Skips if edit check doesn't have _ meaning it's not a dynamic
            continue
        match = False  # Tracks if folder and formOID match is found
        for check_name_cell in check_names_column:  # Iterates through all check names
            if edit_check == check_name_cell.value:
                oid_list = extract_oid(dynamics_sheet[f"C{edit_check_cell.row}"].value)  # Gets OIDs from description
                if len(oid_list) == 0:  # If there are no OIDs, write to output (not sure if this is wanted)
                    output.write(f"C{edit_check_cell.row} does not contain an OID\n")
                    match = True
                    break  # Moves to next edit check
                for oid in oid_list:  # Iterate through each OID in description
                    form_oid = check_steps_sheet[f"H{check_name_cell.row}"].value
                    if "*." in oid:  # Run if there's a wildcard folder
                        split_oid = oid.split(".")
                        if split_oid[1] == form_oid:
                            match = True
                    elif "." in oid:  # Run if there's a folderOID
                        split_oid = oid.split(".")
                        folder_oid = check_steps_sheet[f"G{check_name_cell.row}"].value
                        if split_oid[0] == folder_oid and split_oid[1] == form_oid:
                            match = True
                    else:  # Run if there isn't a folderOID
                        if oid == form_oid:
                            match = True
                            break
                if match:  # Stop looping once match is found
                    break
        if not match:
            output.write(f"Incorrect description in C{edit_check_cell.row}\n")

output.write("\nIncorrect Operators:\n")
operators = {  # Create dictionary of relevant operators
    "!= Blank": "IsNotEmpty",
    "!=": "IsNotEqualTo",
    "<=": "IsLessThanOrEqualTo",
    ">=": "IsGreaterThanOrEqualTo",
    "=": "IsEqualTo",
    "AND": "And",
    "OR": "Or",
    "IsLessThan": "<",
    "IsGreaterThan": ">"
}

# Check that each description has correct set of operators
for edit_check_cell in edit_checks_column:  # Iterates over each edit check cell
    if edit_check_cell.value is None:  # Stops once a blank cell is reached
        break
    edit_check_split = edit_check_cell.value.split("\n")  # Separates multiple checks in one cell
    for edit_check in edit_check_split:  # Iterates over each check inside once cell
        if "_" not in edit_check:  # Skips over non-Dynamics
            continue
        check_functions = list()  # Creates list of CheckFunctions in CheckSteps
        for check_name_cell in check_names_column:  # Adds CheckFunctions to list
            if edit_check == check_name_cell.value and \
                    check_steps_sheet[f"C{check_name_cell.row}"].value is not None:
                check_functions.append(check_steps_sheet[f"C{check_name_cell.row}"].value)
        check_functions.sort()

        description_functions = list()  # Creates list of operators in description
        dynamic_description = dynamics_sheet[f"C{edit_check_cell.row}"].value
        for key in operators:  # Iterates over each operator
            for i in range(dynamic_description.count(key)):  # Adds equivalent CheckFunction to list for each operator
                description_functions.append(operators[key])
            dynamic_description = dynamic_description.replace(key, "")  # Removes operator so no double counts
        description_functions.sort()

        if len(description_functions) == 0:  # Write to output if no operators are found
            output.write(f"C{edit_check_cell.row} is missing operators\n")
        elif description_functions != check_functions:  # Write to output if lists don't match
            output.write(f"{edit_check} in C{edit_check_cell.row} has incorrect operators\n")

output.write("\nIncorrect Trigger Values:\n")
dictionary_sheet = sds_book["DataDictionaryEntries"]  # Get DataDictionaryEntries sheet
data_dictionary_column = dictionary_sheet["A"]  # Get DataDictionary column from DataDictionaryEntries sheet
fields_sheet = sds_book["Fields"]  # Get Fields sheet
fields_column = fields_sheet["B"]

# Check that each description has the correct trigger value
for edit_check_cell in edit_checks_column:  # Iterates over each edit check cell
    if edit_check_cell.value is None:  # Stops once a blank cell is reached
        break
    edit_check_split = edit_check_cell.value.split("\n")  # Separates multiple checks in one cell
    for edit_check in edit_check_split:
        if "_" not in edit_check:  # Skips over non-Dynamics
            continue
        field_oid = ""
        trigger = ""
        dictionary_values = dict()  # Data Dictionary
        for check_name_cell in check_names_column:  # Loop through check names to find match
            if check_name_cell.value == edit_check:
                if check_steps_sheet[f"I{check_name_cell.row}"].value is not None:  # Find fieldOID and dictionary
                    dictionary_values.clear()
                    field_oid = check_steps_sheet[f"I{check_name_cell.row}"].value
                    dictionary_name = ""  # Name of corresponding Data Dictionary
                    for field_cell in fields_column:  # Find corresponding data dictionary
                        if field_cell.value == field_oid and fields_sheet[f"I{field_cell.row}"].value is not None:
                            dictionary_name = fields_sheet[f"I{field_cell.row}"].value  # Store dictionary name

                            for data_dictionary_cell in data_dictionary_column:  # Fill in dictionary values
                                if dictionary_name == data_dictionary_cell.value:
                                    dictionary_values[dictionary_sheet[f"B{data_dictionary_cell.row}"].value] = \
                                        dictionary_sheet[f"D{data_dictionary_cell.row}"].value

                if check_steps_sheet[f"D{check_name_cell.row}"].value is not None:  # Find trigger value
                    trigger = check_steps_sheet[f"D{check_name_cell.row}"].value
                    trigger_string = dictionary_values.get(trigger)
                    dynamic_description = dynamics_sheet[f"C{edit_check_cell.row}"].value
                    if trigger_string not in dynamic_description:
                        output.write(f"C{edit_check_cell.row} has incorrect trigger values\n")

