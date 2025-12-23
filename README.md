# Excel Number to Indian Rupee Text Macro

A powerful VBA macro and User Defined Function (UDF) for Microsoft Excel that converts numerical values into their word representation in the **Indian Rupee (INR)** format. 

*Example: `1,23,456.78` becomes **"Rupees One Lac Twenty Three Thousand Four Hundred Fifty Six and Seventy Eight Paisas Only"***.

---

## 🚀 Features

- **Bulk Conversion**: Select multiple cells and convert them all at once.
- **Separate Destination**: Keeps your original numbers safe by writing the text to a different column.
- **Formula Support**: Use `=CONVERTTOTEXT()` directly in your spreadsheet.
- **Indian Numbering**: Properly handles Thousands, Lacs, and Crores.
- **Decimal Support**: Converts fractions into Paisas automatically.

---

## 🛠️ Step 1: Initial Setup (For Beginners)

If you want this macro to be available in **every** Excel file you open, you should put it in your **Personal Macro Workbook (PERSONAL.XLSB)**. Here is how to create it if you don't have one:

1. **Show the Developer Tab** (if not visible):
   - Right-click any tab on the Ribbon (e.g., Home) and select **Customize the Ribbon**.
   - In the right-hand list, check the box for **Developer** and click **OK**.
2. **Create the PERSONAL.XLSB file**:
   - Go to the **Developer** tab and click **Record Macro**.
   - In the "Store macro in" dropdown, select **Personal Macro Workbook**.
   - Click **OK**, then click any cell in your sheet, and immediately click **Stop Recording**. 
   - *This "tricks" Excel into creating the hidden Personal file for you.*
3. **Open the VBA Editor**:
   - Press `ALT + F11` on your keyboard.
4. **Paste the Code**:
   - In the left-hand pane (Project Explorer), find **VBAProject (PERSONAL.XLSB)**.
   - Right-click the **Modules** folder inside it and select **Insert > Module**.
   - **Rename the Module**: In the Properties window (bottom-left), change the Name from `Module1` to `modNumberToText`. (Important: Do not name it exact same as the function).
   - Copy the entire code from `excelNumberToText.vba` in this repository and paste it into the large white code window.
5. **Save & Exit**:
   - Click the **Save** icon in the VBA editor.
   - Close the VBA window and return to Excel.

---

## 📖 How to Use

### Method 1: Using the Macro (Bulk Processing)
Best for converting hundreds of rows at once.
1. Press `ALT + F8` on your keyboard.
2. Select `PERSONAL.XLSB!ConvertToRupees` and click **Run**.
3. **Step A**: Select the cells containing your numbers.
4. **Step B**: Select the first cell where you want the text to start appearing.
5. The macro will fill the destination cells automatically.

### Method 2: Using the Formula (Dynamic)
Best for individual cells that might change frequently.
1. In any cell, type:
   ```excel
   =PERSONAL.XLSB!CONVERTTOTEXT(A1)
   ```
   *(Where `A1` is the cell with the number).*

> **Tip**: If you find the `PERSONAL.XLSB!` prefix annoying, you can save the code as an **Excel Add-In (.xlam)** instead. Instructions for that can be found in advanced Excel guides.

---

## ⚠️ Troubleshooting #NAME? Error

If Excel shows `#NAME?`, check these three things:
1. **The Prefix**: Ensure you are typing `PERSONAL.XLSB!` before the function name if the code is in your personal workbook.
2. **Module Name**: Make sure your **Module** name is NOT `CONVERTTOTEXT`. Rename it to `modNumberToText` as shown in the setup steps.
3. **Macro Security**: Go to `File > Options > Trust Center > Trust Center Settings > Macro Settings` and ensure macros are enabled.

---

## 📉 Limitations
- **Maximum Value**: Supports up to 99 Crores (`999,999,999.99`).

## 📜 License
MIT License - Feel free to use and modify for personal or commercial projects.
