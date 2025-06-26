# Import Workbook class from openpyxl to create and manipulate Excel files (.xlsx)
from openpyxl import Workbook

# Create a new Excel workbook object
wb = Workbook()

# Get the active worksheet (by default, this is the first sheet in the workbook)
ws = wb.active

# Rename the worksheet title to "Claims" for better context
ws.title = "Claims"

# Define a list of dictionaries, where each dictionary contains one insurance claim's details.
claims  = [
    {
        "INSURED/POLICYHOLDER": "ABIGAIL CARTER",
        "ADDRESS": "SCARBOROUGH, ON M1C 2Z3",
        "INSURER": "ABC INSURANCE",
        "CLAIM #": "PR1923",
        "ADJUSTER/ CLAIM REP": "NOVA CLAIMS",
        "DATE OF INSPECTION": "MARCH 15, 2025",
        "DATE OF LOSS": "MARCH 12, 2025",
        "DATE OF REPORT": "MARCH 18, 2025",
        "TYPE OF LOSS": "WATER DAMAGE",
        "CAUSE OF LOSS": "The loss was caused by a burst pipe on the second floor which led to flooding in multiple rooms.",
        "SCOPE OF WORK": "1. Assess and document all damaged areas.\n2. Extract standing water.\n3. Dry and dehumidify structure.\n4. Restore affected areas.\n5. Dispose of irreparable materials."
    },
    {
        "INSURED/POLICYHOLDER": "MICHAEL LEE",
        "ADDRESS": "MISSISSAUGA, ON L5N 3T4",
        "INSURER": "GUARDIAN CO.",
        "CLAIM #": "PR2145",
        "ADJUSTER/ CLAIM REP": "TOP GUN",
        "DATE OF INSPECTION": "MARCH 21, 2025",
        "DATE OF LOSS": "MARCH 19, 2025",
        "DATE OF REPORT": "MARCH 22, 2025",
        "TYPE OF LOSS": "FIRE",
        "CAUSE OF LOSS": "The loss was caused by a kitchen fire that spread to the living room before being contained.",
        "SCOPE OF WORK": "1. Pack out salvageable items.\n2. Clean soot from surfaces.\n3. Remove damaged cabinetry and appliances.\n4. Restore interior walls and ceiling.\n5. Reinstall fixtures and fittings."
    },
    {
        "INSURED/POLICYHOLDER": "PRIYA SHARMA",
        "ADDRESS": "BRAMPTON, ON L6T 1S9",
        "INSURER": "OMEGA INSURE",
        "CLAIM #": "PR2298",
        "ADJUSTER/ CLAIM REP": "DELTA REP",
        "DATE OF INSPECTION": "APRIL 2, 2025",
        "DATE OF LOSS": "MARCH 30, 2025",
        "DATE OF REPORT": "APRIL 3, 2025",
        "TYPE OF LOSS": "STORM",
        "CAUSE OF LOSS": "Strong winds during a thunderstorm caused a tree to fall on the roof.",
        "SCOPE OF WORK": "1. Remove debris.\n2. Tarp damaged areas.\n3. Replace damaged shingles.\n4. Inspect attic for structural issues.\n5. Restore affected rooms."
    },
    {
        "INSURED/POLICYHOLDER": "OMAR AHMED",
        "ADDRESS": "ETOBICOKE, ON M9C 1W4",
        "INSURER": "XYZ COVERAGE",
        "CLAIM #": "PR3010",
        "ADJUSTER/ CLAIM REP": "FALCON HANDLER",
        "DATE OF INSPECTION": "FEBRUARY 28, 2025",
        "DATE OF LOSS": "FEBRUARY 26, 2025",
        "DATE OF REPORT": "MARCH 1, 2025",
        "TYPE OF LOSS": "THEFT",
        "CAUSE OF LOSS": "Theft occurred while the insured was on vacation. Multiple valuables and electronics were stolen.",
        "SCOPE OF WORK": "1. Inventory missing items.\n2. Assess property damage.\n3. Replace broken windows/locks.\n4. Secure the property.\n5. Coordinate with local authorities."
    },
    {
        "INSURED/POLICYHOLDER": "EMILY ZHANG",
        "ADDRESS": "NORTH YORK, ON M2J 5C3",
        "INSURER": "NORTHSHORE INSURANCE",
        "CLAIM #": "PR3123",
        "ADJUSTER/ CLAIM REP": "EAGLE EYE",
        "DATE OF INSPECTION": "APRIL 4, 2025",
        "DATE OF LOSS": "MARCH 31, 2025",
        "DATE OF REPORT": "APRIL 5, 2025",
        "TYPE OF LOSS": "VANDALISM",
        "CAUSE OF LOSS": "Graffiti and broken exterior windows were discovered after a break-in attempt.",
        "SCOPE OF WORK": "1. Remove graffiti.\n2. Replace broken glass.\n3. Repair window frames.\n4. Repaint walls.\n5. Install motion lighting."
    },
    {
        "INSURED/POLICYHOLDER": "JORDAN WILLIAMS",
        "ADDRESS": "OAKVILLE, ON L6M 2W2",
        "INSURER": "GUARDIAN CO.",
        "CLAIM #": "PR3489",
        "ADJUSTER/ CLAIM REP": "NOVA CLAIMS",
        "DATE OF INSPECTION": "MARCH 25, 2025",
        "DATE OF LOSS": "MARCH 22, 2025",
        "DATE OF REPORT": "MARCH 26, 2025",
        "TYPE OF LOSS": "FIRE",
        "CAUSE OF LOSS": "The fire started in the garage due to an electrical fault and spread to adjoining rooms.",
        "SCOPE OF WORK": "1. Clean and deodorize affected area.\n2. Restore charred framing.\n3. Replace drywall.\n4. Rewire electrical connections.\n5. Paint and finish surfaces."
    },
    {
        "INSURED/POLICYHOLDER": "ANITA DESAI",
        "ADDRESS": "SCARBOROUGH, ON M1V 3P2",
        "INSURER": "OMEGA INSURE",
        "CLAIM #": "PR3590",
        "ADJUSTER/ CLAIM REP": "DELTA REP",
        "DATE OF INSPECTION": "MARCH 27, 2025",
        "DATE OF LOSS": "MARCH 25, 2025",
        "DATE OF REPORT": "MARCH 28, 2025",
        "TYPE OF LOSS": "WATER DAMAGE",
        "CAUSE OF LOSS": "A leak in the upstairs bathroom resulted in extensive ceiling damage below.",
        "SCOPE OF WORK": "1. Stop the source of leak.\n2. Remove wet insulation.\n3. Replace ceiling panels.\n4. Sanitize affected zones.\n5. Dry and repaint surfaces."
    },
    {
        "INSURED/POLICYHOLDER": "LUCAS BENOIT",
        "ADDRESS": "TORONTO, ON M4P 1Z2",
        "INSURER": "NORTHSHORE INSURANCE",
        "CLAIM #": "PR3766",
        "ADJUSTER/ CLAIM REP": "FALCON HANDLER",
        "DATE OF INSPECTION": "MARCH 12, 2025",
        "DATE OF LOSS": "MARCH 9, 2025",
        "DATE OF REPORT": "MARCH 13, 2025",
        "TYPE OF LOSS": "STORM",
        "CAUSE OF LOSS": "Heavy rains flooded the basement and damaged electrical outlets.",
        "SCOPE OF WORK": "1. Pump out flood water.\n2. Replace affected wiring.\n3. Inspect mold presence.\n4. Install sump pump.\n5. Replace carpeting."
    },
    {
        "INSURED/POLICYHOLDER": "SANDRA KIM",
        "ADDRESS": "RICHMOND HILL, ON L4B 4M6",
        "INSURER": "XYZ COVERAGE",
        "CLAIM #": "PR3844",
        "ADJUSTER/ CLAIM REP": "EAGLE EYE",
        "DATE OF INSPECTION": "APRIL 5, 2025",
        "DATE OF LOSS": "APRIL 2, 2025",
        "DATE OF REPORT": "APRIL 6, 2025",
        "TYPE OF LOSS": "VANDALISM",
        "CAUSE OF LOSS": "Unknown individuals defaced the front door and damaged the mailbox.",
        "SCOPE OF WORK": "1. Document and photograph damage.\n2. Replace mailbox.\n3. Repair door and repaint.\n4. Install surveillance camera.\n5. Notify local authorities."
    },
    {
        "INSURED/POLICYHOLDER": "KEVIN O'NEILL",
        "ADDRESS": "VAUGHAN, ON L6A 1R2",
        "INSURER": "ABC INSURANCE",
        "CLAIM #": "PR3975",
        "ADJUSTER/ CLAIM REP": "TOP GUN",
        "DATE OF INSPECTION": "MARCH 18, 2025",
        "DATE OF LOSS": "MARCH 15, 2025",
        "DATE OF REPORT": "MARCH 20, 2025",
        "TYPE OF LOSS": "THEFT",
        "CAUSE OF LOSS": "Forced entry through the back door resulted in stolen electronics and cash.",
        "SCOPE OF WORK": "1. Replace broken lock.\n2. Inventory stolen items.\n3. Restore damaged doorframe.\n4. Reinstall security system.\n5. Submit police report."
    }
]

# Extract column headers from the first claim's keys
headers = list(claims[0].keys())

# Write the headers as the first row in the Excel sheet
ws.append(headers)

# Loop over each claim dictionary in the list
for claim in claims:
    # Create a row by accessing each value in the same order as the headers
    row = [claim[key] for key in headers]
    # Append the row to the worksheet
    ws.append(row)

# Save the workbook as an Excel file
# Note: openpyxl creates .xlsx files, not .xls (which is an older Excel format)
wb.save("data.xlsx")  # Fixed from "data.xls" to correct format
print("file created successfully")