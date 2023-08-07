# CopyData Macro for Excel

## Description

The "CopyData" macro is a versatile Excel automation tool designed to streamline the process of copying data from one workbook to another. It is particularly useful when you have two open workbooks, the "Host" workbook (the workbook containing the macro or the button) and the "Data" workbook (the workbook from which data needs to be copied). The macro automatically finds the "Data" workbook and transfers matching data from the "Data" workbook to the "Host" workbook, based on column headers.

## How it Works

1. The macro first identifies the "Host" workbook (the workbook containing the macro or the button) and searches through all open workbooks to find the "Data" workbook.
2. It then identifies the worksheets in both workbooks.
3. The macro proceeds to find the last used row and column numbers in both the "Host" and "Data" worksheets.
4. It loops through the column headers of both worksheets to find matching column headers (case-insensitive comparison).
5. When a match is found, the macro copies the data from the corresponding column in the "Data" worksheet to the matching column in the "Host" worksheet. It ensures that the data size in both columns is aligned to avoid any data mismatch. Rest of the columns in the "Host" workbook will be left unchanged.
6. The process continues until all matching columns are copied.
7. After the data is copied, the macro displays a message box confirming the successful data transfer.

## Usage

The "CopyData" macro can be used in various scenarios, such as merging data from different files, consolidating information from multiple sheets, or updating a master workbook with data from various sources. The macro saves valuable time and reduces the chances of manual errors during data copying.

## Dependencies

The macro requires Microsoft Excel to run and is compatible with various versions of Excel.

## Instructions for Use

1. Open the "Host" workbook (the workbook containing the macro or the button).
2. Open the "Data" workbook (the workbook from which data needs to be copied).
3. Enable macros in Excel, if not already enabled.
4. Run the "CopyData" macro from the "Host" workbook or click the button placed in the "Host" workbook.
5. The macro will automatically identify the "Data" workbook and transfer matching data to the "Host" workbook. The macro will leave any additional columns in the "Host" workbook unchanged.
6. A message box will display upon successful data transfer.

**Important Note:** It is advisable to back up your workbooks before using the macro, especially if you are working with critical data.

Feel free to customize and upload this "README.md" file to your GitHub repository to document the enhanced functionality of the "CopyData" macro. If you have any other requests or need further assistance, please don't hesitate to ask!
