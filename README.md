# Excel Number to Indian Rupee Text Macro

A powerful VBA macro and User Defined Function (UDF) for Microsoft Excel that converts numerical values into their word representation in the **Indian Rupee (INR)** format (e.g., 1,23,456.78 becomes "Rupees One Lac Twenty Three Thousand Four Hundred Fifty Six and Seventy Eight Paisas Only").

## Features

- **Bulk Conversion**: Select multiple cells at once and convert them to text.
- **Destination Control**: Choose a separate range for the output text to preserve your original numerical data.
- **Global Function**: Includes the `=CONVERTTOTEXT()` function for use directly in Excel formulas.
- **Indian Numbering System**: Tailored for Lacs and Crores.
- **Decimal Support**: Handles Paisas automatically.

## How to Install

1. Open your Excel Workbook.
2. Press `ALT + F11` to open the VBA Editor.
3. Click `Insert` > `Module`.
4. Copy the contents of `excelNumberToText.vba` from this repository and paste it into the module.
5. Close the VBA Editor.
6. Save your workbook as an **Excel Macro-Enabled Workbook (.xlsm)**.

## How to Use

### Method 1: Using the Macro (Bulk Conversion)
1. Press `ALT + F8` and select `ConvertToRupees`.
2. Select the cells containing the numbers you want to convert.
3. Select the destination cells where you want the text to appear.
4. The macro will validate the range sizes and perform the conversion.

### Method 2: Using the Formula
1. In any cell, type:
   ```excel
   =CONVERTTOTEXT(A1)
   ```
   (Where `A1` is the cell containing the number).

## Limitations
- Maximum supported value: 99 Crores (999,999,999.99).

## License
MIT License
