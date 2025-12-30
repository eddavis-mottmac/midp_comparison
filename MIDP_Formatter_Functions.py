def format_midp_df(midp):
    # Dictionary with column mappings
    midp = midp[midp["Publish Row to TW"] == "TRUE"]
    # Change the value of 'DCO Deliverable' column to 'Y' if true, and 'N' if false or blank
    midp['DCO Submission Document'] = midp['DCO Submission Document'].apply(lambda x: 'Y' if x == "TRUE" else 'N')
    
    print(midp.shape)

    columns_required = {
        "Information Container": "Information Container",
        "Information Container Title / Description": "Information Container Title",
        "ID number": "P6 Activity ID (Please Review)",
        "Project Milestone": "LWR - Phase",
        "Comment": "Comments",
        "ID": "",
        "LOA": "",
        "Status Code": "⚡ Document Status",
        "Revision": "⚡ Last Published Revision",
        "Issue Date (Planned)": "Planned Issue Date",
        "Issue Date (Actual)": "⚡ Published Date",
        "Authorised by TW Service manager": "",
        "Created by": "Document Workstream (Owning team)",
        "Y/N": "DCO Submission Document",
        "Y/N.": "",
        "File Extension": "Source File Extension (Please Review)",
        "Security Reference": "Security Reference",
        "Project": "Project",
        "Functional Breakdown": "Functional Breakdown",
        "Spatial Breakdown": "Spatial Breakdown",
        "Document Type": "Document Type",
        "Discipline": "Discipline"
    }

    # Create a list of all the values in the dictionary
    values_list = list(columns_required.values())
    keys_list = list(columns_required.keys())

    # Remove blank values from the list
    values_list = [value for value in values_list if value]


    # Create a new dictionary with keys and values swapped
    columns_required_swapped = {v: k for k, v in columns_required.items() if v}
    print(columns_required_swapped)
    # Rename columns in the DataFrame according to the new dictionary
    midp.rename(columns=columns_required_swapped, inplace=True)

    # Drop columns from midp DataFrame if not in values of columns_required
    midp = midp[[col for col in midp.columns if col in keys_list]]

    # Add missing columns with blank values
    for col in keys_list:
        if col not in midp.columns:
            midp[col] = ""

    # Fill any remaining NaN values with empty strings
    midp.fillna('', inplace=True)


    # Reorder columns as per original_keys
    midp = midp[keys_list]

    
    
    return midp

