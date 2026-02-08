"""PowerPoint VBA CLI module."""

from vba_edit.office_cli import create_office_main

# Create the main function for PowerPoint
main = create_office_main("powerpoint")

if __name__ == "__main__":
    main()
