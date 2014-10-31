# Varrow's Accounting Department Portal
## Allows them to:
	* Upload different types of employee expenses in the form of excel spreadsheets to Intacct web server
	* Automatically create bills in Intacct software
	* Search for specific Salesforce user accounts
	* Create Opportunities, Opportunity Line Items, Sales Invoices, and Sales Invoice Line Items in Salesforce with the click of one button
## Things about the code:
	* Written in JavaScript using NodeJS to handle the web server things
	* JSForce to connect to Salesforce and for authentication. Also allows users to create Salesforce objects with proper permissions
	* Uses Jade Templating Engine to pass data directly from the code to the currently displayed page
	* JS-XLSX to parse through excel worksheets and grab pertinent data
	* XML-Builder to create an XML contianing worksheet data, so it can be shipped off to Intacct
## To use:
	* Log in using your Varrow Salesforce info
	* Click on "Push Concur Employee Expenses" to upload regular expenses
	* Click on "Push Amex Expenses" to upload Amex exepnses
	* Click "Search For Accounts" to search for an account and create Salesforce objects based on that account

## TODO:
    * Form Validation
    * Modularization of code
    * Update Amex expenses for Concur Breeze
