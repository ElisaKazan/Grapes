from openpyxl import load_workbook
import sys

# To run script:
# python grapes.py <name-of-file>

def main():
    # Get name of excel document
    if len(sys.argv) < 2:
        # Error message & exit
        print("Usage: python grapes.py <excel-file-name>")
        sys.exit()
    excel_name = sys.argv[1]
    print("Document name: {}".format(excel_name))

    # Open excel document
    excel_workbook = load_workbook(excel_name)

    # Create grades dictionary
    grades_dict = {}

    # Loop over sheets
    for sheet in excel_workbook:
        print("Sheet - {}".format(sheet.title))

        # Loop over rows (starting at 3)
        for row in sheet.iter_rows(min_row=3, min_col=1, max_col=4):
            percentage = get_percentage(row[3].value)
            print("\tAssignment {} - {:.3f}%".format(row[0].value, float(row[2].value) * percentage))
            #for cell in row:
            #    print(cell.value)



        # Get Weight percentage (ex: 4)

        # Get Grade Received (ex: 90)

        # TODO:check formatting

        # Find Grade Contribution (ex: 4 * (90/100))

        # Print all values

        # Save Grade Contribution

    # Save Worksheet

    # Print Happy Message

def get_percentage(grade):
    if isinstance(grade, (int,float)):
        if grade < 1:
            return float(grade)
        elif grade < 100:
            return float(grade)/100
        else:
            print("Error: Not a valid grade")
            sys.exit()
    elif '/' in grade:
        nums = grade.split('/')
        return float(nums[0])/float(nums[1])
    else:
        print("Error: Something went wrong")


if __name__ == '__main__':
    # executes only if run as script
    main()