"""Word VBA CLI module."""

from vba_edit.office_cli import create_office_main

# Create the main function for Word
main = create_office_main("word")

if __name__ == "__main__":
    main()
