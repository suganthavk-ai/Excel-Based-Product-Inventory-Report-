# Excel-Based Product Inventory Report
This Dataset for practicing data cleaning, data transformation, conditional formatting, and inventory analysis in MS Excel.

## Description
This dataset contains information about various products from different countries, including their manufacturing details, brand information, pricing, inventory quantity, and product categories.

## Project Objective
The project aims to:
•	Perform data cleaning and preprocessing. 
•	Handle missing values and inconsistencies. 
•	Standardize text and date formats. 
•	Apply conditional formatting rules. 
•	Analyze product pricing and inventory levels. 
•	Generate reports and dashboards in Excel. 

## Getting Started
### Dependencies
The following tools and software were used in this project:
• OS Version
• MS Office latest version
• Knowledge of Excel formula

## Dataset Description
|Column Name | Description|
|-------| --------------|
|Manufacturing Date | Date when the product was manufactured |
|Country Code | Country identifier where the product is associated |
|Product ID | Unique identifier for each product |
|Product Brand | Product brand or model information |
|Product Name | Name of the product |
|Brand Name | Manufacturer or brand of the product |
|Price ($) | Selling price of the product |
|Quantity | Available inventory quantity | 
|Category | Product classification (Electronics, Fashion, Kitchen, Outdoor, Accessories, etc.) |

## Installation
1. Download the dataset file.
2. Save a copy of the file on your PC or laptop.
3. Open the file using Microsoft Excel.
4. If a Protected View message appears, click Enable Editing.
5. Review the dataset and explore the product inventory and pricing details.
6. Use Excel features such as:
	Sorting and Filtering
	Find and Replace 
	Conditional Formatting 
	Text Functions (PROPER, TRIM, CLEAN) 
	IF Function 
	AVERAGE Function 
	Cell Formatting (DD-MM-YYYY format and format currency values)
	Data Validation 
	Tables 
	Basic Data Cleaning Techniques 

## Executing program
### Step to run the program
•	Open the Excel-Based Product Inventory Report Excel file.
•	Ensure that editing is enabled in Microsoft Excel.
•	Verify that all dataset columns are properly formatted.

## Some Basic Aggregate Functions             
### AVERAGE Function
(=AVERAGE(D2:D35))
### ISBLANK Function
(=IF(ISBLANK(D2),AVERAGE($D$2:$D35),D2)
(=IF(ISBLANK(F2),”Unknown”,F2)
### CLEANING Function
(=PROPER(TRIM(CLEAN(B2)))
### SPLIT Function
(=TRIM(CLEAN(A2)), (=LEFT((A2),6), (=RIGHT(A2),2)
### MERGE Function
(=CONCATENATE(B2,” – “,C2))

# Help

## Tips to Avoid Errors
•	Always verify cell references before applying formulas.
•	Use proper data formatting.
•	Check for empty or incorrect values in the dataset.

## Authors

Sugantha B

## Version History
### 1.0       
•	Initial Release
### 1.1 
•	Data Update
### 1.2 
•	Data Cleaning and Preprocessing
### 1.3 
•	Formula Implementation
### 1.4
•	Data Analystis

 ## License
This project is created for Educational and Assignment Purposes only.
## Acknowledgments
•	MS Office
•	GitHub
•	Online Educational resources 
This project helped improve practical knowledge in:
•	Perform data cleaning and preprocessing. 
•	Handle missing values and inconsistencies. 
•	Standardize text and date formats. 
•	Apply conditional formatting rules. 
•	Analyze product pricing and inventory levels. 
•	Generate reports and dashboards in Excel.

Their support and resources contributed greatly to the successful completion of the Excel-Based Product Inventory Report project.



