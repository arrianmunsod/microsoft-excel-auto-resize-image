# Microsoft Excel – Automated Image Resizing to Cell Dimensions

Automatically resize and center pictures within Excel cells or merged cells while maintaining their original aspect ratios.

## Overview

This VBA automation resizes selected images to fit within their corresponding Excel cells, including merged cells.

The macro is designed to be stored in Excel's **Personal Macro Workbook (PERSONAL.XLSB)**, allowing you to use the same functionality across multiple Excel workbooks without having to install the macro individually in each file.

## Features

- Automatically resizes pictures to fit within cell dimensions.
- Supports regular and merged cells.
- Maintains original image aspect ratios.
- Automatically centers pictures within cells.
- Supports customizable margins.
- Supports multiple selected pictures.
- One-click execution through the Quick Access Toolbar (QAT).
- Works across different Excel workbooks.
- No need to save individual workbooks as macro-enabled files (`.xlsm`).
- Automatically loads the macro whenever Excel starts.

## Repository Contents

| File | Description |
|---|---|
| `ExcelAutoResizeImage.bas` | VBA module containing the image-resizing automation script. |

## Installation Guide

### 1. Open Microsoft Excel

Open Microsoft Excel using a blank workbook or any existing workbook.

### 2. Create the Personal Macro Workbook (PERSONAL.XLSB)

If you haven't created a Personal Macro Workbook before, follow these instructions:

1. Navigate to **View → Macros → Record Macro**.
2. Enter any temporary macro name, such as `CreatePersonalWorkbook`.
3. Under **Store macro in**, select **Personal Macro Workbook**.
4. Click **OK** to start recording.
5. Navigate to **View → Macros → Stop Recording**.

Excel will automatically create the `PERSONAL.XLSB` file.

**Important:** You only need to perform this step once. If you already have a Personal Macro Workbook, skip this section.

### 3. Open the VBA Editor

- Press `ALT + F11` to open the Visual Basic for Applications (VBA) Editor.
- Locate **VBAProject (PERSONAL.XLSB)** in the Project Explorer.

If the Project Explorer is not visible, press `CTRL + R`.

### 4. Import the VBA Module

1. Expand **VBAProject (PERSONAL.XLSB)** in the Project Explorer.
2. Right-click the project name.
3. Select **Import File...**.
4. Locate and select `ExcelAutoResizeImage.bas` from this repository.
5. Click **Open**.

The imported module should now appear under the **Modules** folder.

**Alternative:** You can also open the `.bas` file using a text editor, copy its contents, and paste the code into a newly created VBA module.

### 5. Save the Personal Macro Workbook

1. Press `CTRL + S` while inside the VBA Editor, or return to Excel and save the Personal Macro Workbook.
2. Close Excel.
3. If Excel asks whether you want to save changes to the Personal Macro Workbook, select **Save**.

**Important:** Make sure you save `PERSONAL.XLSB`. Otherwise, your newly imported macro may be lost.

You do not need to save your individual Excel workbooks as `.xlsm`.

### 6. Add the Macro to the Quick Access Toolbar

1. Open Excel.
2. Navigate to **File → Options**.
3. Select **Quick Access Toolbar**.
4. Under **Choose commands from**, select **Macros**.
5. Locate the macro:

   `PERSONAL.XLSB!FitImageWithMargin`

6. Select it and click **Add >>**.
7. Optionally, click **Modify** to customize the button icon and display name.
8. Click **OK**.

You should now see your image-resizing button in the Quick Access Toolbar.

## How to Use the Macro

1. Open any Excel workbook.
2. Insert your pictures into the worksheet.
3. Position each picture over its intended cell.
4. Select one or multiple pictures.
5. Click the **Fit Image** button in the Quick Access Toolbar.

The macro will automatically:

- Identify the cell underneath each picture.
- Detect merged cells.
- Calculate the available cell dimensions.
- Resize pictures while maintaining their original aspect ratios.
- Apply the specified margins.
- Center pictures within their corresponding cells.
- Set picture placement to **Move and Size with Cells**.

## Customization

You can customize the spacing between the picture and its corresponding cell boundaries by modifying the `Margin` variable in the VBA module.

Default setting:

`Margin = 2`

Excel measures shape and cell dimensions in points rather than pixels. Therefore, the default value of `2` represents a 2-point margin on each side.

You can increase or decrease this value depending on your preferred spacing.

## Important Notes

- **PERSONAL.XLSB is stored locally:** The macro is available on the computer where you installed it. If you use another computer, you must transfer or recreate your Personal Macro Workbook.
- **No individual XLSM files required:** You can continue saving your regular Excel workbooks as `.xlsx`.
- **Macro security:** Excel must allow macros to run for the functionality to work.
- **Save changes:** Whenever you modify the VBA code, save the Personal Macro Workbook.
- **Workbook compatibility:** The macro works with regular and merged cells, provided the picture's top-left anchor is positioned over the intended cell.

## Final Result

You now have a reusable Excel image-resizing tool installed directly in your Personal Macro Workbook.

Instead of copying VBA code into every new workbook, you can simply open Excel, select your pictures, and click your Quick Access Toolbar button.

## Development Note

AI-assisted development. The code was reviewed and tested by the author.

---

<p align="center">
  <a href="https://github.com/arrianmunsod">
    <strong>← BACK TO ARRIAN'S GITHUB PROFILE</strong>
  </a>
  <br>
  <sub>Explore more projects, scripts, and technical documentation.</sub>
</p>
