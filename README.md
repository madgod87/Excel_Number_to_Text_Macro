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

## Troubleshooting #NAME? Error

If Excel does not recognize the formula, check the following:

### 1. The Prefix Requirement
If you have pasted the code into your **PERSONAL.XLSB** (Personal Macro Workbook), you must use the prefix:
```excel
=PERSONAL.XLSB!CONVERTTOTEXT(A1)
```

### 2. Module Naming Conflict
**IMPORTANT**: Do not name the Module the same as the Function.
- In the VBA Editor, if your module is named `CONVERTTOTEXT`, rename it to `modNumberToText`.

### 3. Truly Global Usage (No Prefix)
To use `=CONVERTTOTEXT()` without any prefix in every workbook:
1. Copy the code into a new Excel file.
2. Save the file as an **Excel Add-In (.xlam)**.
3. Enable the Add-in via `Developer` > `Excel Add-ins`.

## Limitations
- Maximum supported value: 99 Crores (999,999,999.99).

## License
MIT License
