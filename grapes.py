from openpyxl import load_workbook
import sys

# python grapes.py <name-of-file>

def main():
    # Get name of excel document
    if len(sys.argv) < 2:
        # Error message & exit
        print("Usage: python grapes.py <excel-file-name>")
        sys.exit()
    excel_doc = sys.argv[1]
    print("Document name: {}".format(excel_doc))
    # Open excel document

    # Create grades dictionary

    # Loop over rows (starting at 3)

        # Get Weight percentage (ex: 4)

        # Get Grade Received (ex: 90)

        # TODO:check formatting

        # Find Grade Contribution (ex: 4 * (90/100))

        # Print all values

        # Save Grade Contribution

    # Save Worksheet

    # Print Happy Message

if __name__ == '__main__':
    # executes only if run as script
    main()