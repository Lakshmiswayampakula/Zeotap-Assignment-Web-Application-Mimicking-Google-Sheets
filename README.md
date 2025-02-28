#SheetMaster - Google SpreadSheets Clone 📊
A full-featured spreadsheet web application mimicking the core functionalities and interface of Google Sheets, with powerful mathematical operations and data quality tools.
🚀 Overview
SheetMaster is a web-based spreadsheet application that closely replicates the Google Sheets experience. This application provides users with a familiar interface and essential spreadsheet functionality including:
1.	Complete spreadsheet interface matching Google Sheets' layout and interaction model
2.	Mathematical functions for numerical data processing
3.	Data quality tools for text manipulation and data cleaning
4.	Rich cell formatting options for professional-looking spreadsheets
5.	Intuitive drag functionality for efficient spreadsheet management
6.	Image integration capabilities for enhanced visual spreadsheets 
7.	Spreadsheets Download functionality for offline access and sharing


Spreadsheet Interface:
✅ Google Sheets-like UI - Toolbar, formula bar, cell grid, and sheet tabs
✅ Drag & Drop Functionality - Seamlessly move content between cells
✅ Cell Dependencies - Automatic recalculation when referenced cells change
✅ Text Formatting - Bold, italic, font size, and color options
✅ Row & Column Management - Add, delete, and resize rows and columns
Mathematical Functions:
✅ SUM - Calculate the total of selected numeric values
✅ AVERAGE - Find the mean value of a range
✅ MAX - Identify the highest value in a selection
✅ MIN - Identify the lowest value in a selection
✅ COUNT - Tally the number of cells containing numerical values
Data Quality Functions:
✅ TRIM - Remove extra whitespace from text
✅ UPPER - Convert text to uppercase
✅ LOWER - Convert text to lowercase
✅ REMOVE_DUPLICATES - Eliminate duplicate rows from a range
✅ FIND_AND_REPLACE - Search and replace text within selected cells
Data Entry & Validation:
✅ Multi-format Data Entry - Support for numbers, text, and dates
✅ Data Validation - Basic type checking for cell contents
✅ Formula Support - Enter and evaluate spreadsheet formulas
 

🎯 Usage Guide
Basic Spreadsheet Navigation:
•	Click on any cell to select it
•	Use arrow keys or Tab/Shift+Tab to navigate between cells
•	Double-click a cell to edit its contents
•	Use the formula bar to enter or modify cell content

Working with Functions:
•	Select a destination cell
•	Type "=" followed by a function name (e.g., "=SUM(")
•	Select a range of cells or enter cell references
•	Close the parentheses and press Enter

Mathematical Function Examples:
•	=SUM(A1:A10) - Adds all values in cells A1 through A10
•	=AVERAGE(B5:B15) - Calculates the average of values in range B5:B15
•	=MAX(C1:D10) - Returns the maximum value in the rectangular range C1:D10
•	=MIN(A1,B1,C1) - Returns the minimum value among cells A1, B1, and C1
•	=COUNT(A1:A20) - Counts cells containing numbers in range A1:A20

Data Quality Function Examples:
•	=TRIM(A1) - Removes leading/trailing spaces from text in A1
•	=UPPER(B5) - Converts text in cell B5 to uppercase
•	=LOWER(C10) - Converts text in cell C10 to lowercase

🧪 Testing
Function Testing:
	Enter sample data across multiple cells
	Apply various functions to test their behavior
	Verify results match expected calculations
	Test edge cases (empty cells, text in numeric functions, etc.)
User Interface Testing:
	Test drag functionality for selection and content
	Verify cell dependencies update correctly
	Confirm formatting options apply as expected
	Test row/column additions and deletions


🚧 Roadmap & Future Enhancements
🔜 Advanced Functions - Statistical, logical, and lookup functions
🔜 Data Visualization - Charts and conditional formatting
🔜 Collaboration Features - Real-time multi-user editing
🔜 Import/Export - Support for Excel, CSV, and PDF formats
🔜 Mobile Optimization - Responsive design for smaller screens

🤝 Contributing
Contributions to enhance SheetMaster are welcome!
Fork this repository
	Create a feature branch (git checkout -b feature/amazing-feature)
	Commit your changes (git commit -m 'Add amazing feature')
	Push to the branch (git push origin feature/amazing-feature)
	Open a Pull Request
