#Mohammed Kassem
#RF Optimization Manager
#
import xml.etree.ElementTree as ET
import xlsxwriter

# File path
file_path = 'your path to nokia 2G dump'

# Parse the XML file
try:
    tree = ET.parse(file_path)
    root = tree.getroot()
except FileNotFoundError:
    print(f"File not found: {file_path}")
    exit()
except ET.ParseError as e:
    print(f"Error parsing XML: {e}")
    exit()

# Initialize Excel workbook
workbook = xlsxwriter.Workbook('output_with_distname1.xlsx')

# Dictionary to track worksheets and their current row
sheets = {}

# Process the XML file
for child in root:
    for step_child in child:
        obj_class = step_child.get('class')  # Object class (e.g., BSC, BTS, etc.)
        dist_name = step_child.get('distName')  # Get the distName attribute
        if not obj_class or not dist_name:
            continue

        # Parse the distName to extract components (e.g., BSC, BCF, BTS)
        dist_parts = {part.split('-')[0]: part.split('-')[1] for part in dist_name.split('/') if '-' in part}

        # Create a new sheet for the object class if not already created
        if obj_class not in sheets:
            sheets[obj_class] = {
                "worksheet": workbook.add_worksheet(obj_class),
                "current_row": 1,  # Start after header row
                "headers": {},  # To track column headers dynamically
            }

        sheet = sheets[obj_class]
        worksheet = sheet["worksheet"]

        # Ensure 'distName' components are added as headers in the first columns
        for idx, key in enumerate(dist_parts.keys()):
            if key not in sheet["headers"]:
                col = len(sheet["headers"])
                sheet["headers"][key] = col
                worksheet.write(0, col, key)

        # Write the distName components to the first columns
        current_row = sheet["current_row"]
        for key, value in dist_parts.items():
            col = sheet["headers"][key]
            worksheet.write(current_row, col, value)

        # Write other attributes and lists in step_child
        for item in step_child:
            name = item.get('name')
            value = item.text

            if name:
                # Handle list items separately
                if item.tag == "list":
                    # Extract values from <p> tags inside the list
                    list_values = [p.text for p in item.findall('p')]
                    value = ', '.join(list_values)  # Join values as a single string

                # Dynamically assign a column for each unique header
                if name not in sheet["headers"]:
                    col = len(sheet["headers"])
                    sheet["headers"][name] = col
                    worksheet.write(0, col, name)  # Write header to first row

                col = sheet["headers"][name]
                worksheet.write(current_row, col, value)

        # Increment the current row for the sheet
        sheet["current_row"] += 1

# Save the workbook
workbook.close()
print("Data successfully written to 'output_with_distname1.xlsx'.")
