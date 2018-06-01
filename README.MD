# JPK Creator

## What is this?
This is VBA macro made for Excel tables with invoices to create JPK_VAT file. 

## ... and what is JPK_VAT file?
JPK_VAT file is a monthly summary of all invoices in given month. This is required by Polish government from every company in Poland. 

## What do I need to use this macro?

You need:
* MS Excel 2007 or later
* Enabled macros in Excel
* Know how to import macro file in excel
* Worksheet with proper data tables

This macro need 2 sheets in workbook to work. 

## The Main table
Main table must be formatted as a table and be in first sheet in workbook (leftmost card). Also it has to contain columns like this:

|  Sales tax code  |  Invoice  |  Address  |  Tax exempt number  |  Customer/Vendor  |  Document date  |  Date  |  Amount origin  |  Sales tax amount  |
|------------------|-----------|-----------|---------------------|-------------------|-----------------|--------|-----------------|--------------------|

Table can have other columns but it has to have mentioned column names. Order has no matter

## The Mapping table
Second required table is 3-column table with tax code mappings. This table must be in sheet named "Mapping" and has to contain 3 columns. Just like this: 

|  Tax code  |  Amount  |  column K_**  |
|------------|----------|---------------|


Headings of table must be in first row of sheet. First column contains tax codes, second ("Amount") can has "Origin" or "Sales tax amount" strings. By this values macro recognise which value should take from main table (Amount origin or Sales tax amount) and the third column specifies where to put that value. "column K_**" column contains numbers of JPK columns to put value into.

## Working scheme

 This macro:

1. Checks if tables are ok
2. Checks if main table is not empty
3. Makes pivot table from main table
4. Makes sheet with JPK_VAT
5. Reads pivot into data array
6. Maps data array into sales and purcharses arrays ready to write into the JPK_VAT sheet
7. Count sums
8. Writes data into JPK_VAT sheet

## Why pivot table? 
Creating pivot table is fast and very good for cleaning your input data. It can reduce number of your entries even by 10 times. It has significant influecne on performance of this script

## And what about performance?
In real life example with 200k entries in main table it can finish job in about 25 seconds. I have tested script without using pivot table and the same example took about 85 seconds to finish.

## Is this macro finished?
Unfortunately no. This script creates only ready to export sheet in excel. Next versions will have feature to export data to CSV and XML files

### License
This project is made under MIT license