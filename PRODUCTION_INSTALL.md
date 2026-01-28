# FillGaps Add-in Production Installation Guide

This guide covers how to install and use the FillGaps Excel add-in from GitHub Pages on your Mac.

## Prerequisites

- Microsoft Excel for Mac (Microsoft 365 or Office 2019+)
- Internet connection to load the add-in from GitHub Pages

## Installation Steps

### Step 1: Download the Production Manifest

Download the production manifest file from:

```
https://mikedtcm22.github.io/excel_conditional_fill/manifest-production.xml
```

Or download directly from the GitHub repository:

```
https://raw.githubusercontent.com/mikedtcm22/excel_conditional_fill/main/manifest-production.xml
```

Save the file to a known location (e.g., Downloads folder).

### Step 2: Create the Sideload Folder

Open Terminal and run:

```bash
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
```

### Step 3: Copy the Manifest

Copy the downloaded manifest to the sideload folder:

```bash
cp ~/Downloads/manifest-production.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/FillGaps-manifest.xml
```

Note: The manifest must be renamed to end with `-manifest.xml` for Excel to recognize it.

### Step 4: Restart Excel

Completely quit Excel (Cmd+Q) and reopen it.

### Step 5: Insert the Add-in

1. Open any Excel workbook
2. Go to **Insert** menu
3. Click **My Add-ins**
4. Under **Developer Add-ins**, find **FillGaps**
5. Click to insert the add-in

The FillGaps group should now appear in your Home ribbon tab.

## Using FillGaps

### Access Methods

FillGaps provides three ways to trigger fill operations:

#### 1. Ribbon Buttons
- Located in the **Home** tab under the **FillGaps** group
- Click **Fill Blanks** or **Fill Errors**

#### 2. Keyboard Shortcuts (Mac)
| Action | Shortcut |
|--------|----------|
| Fill Blanks | **Cmd+Shift+B** |
| Fill Errors | **Cmd+Shift+E** |

#### 3. Context Menu (Right-Click)
1. Select a range of cells
2. Right-click to open the context menu
3. Click **Fill Blanks** or **Fill Errors**

### Usage Workflow

1. **Position your template formula**: The active cell (the cell with the blue border) should contain the formula you want to copy
2. **Select your target range**: Highlight the range where you want formulas filled (this can include the template cell)
3. **Trigger the fill operation** using any of the three access methods above
4. Blank or error cells in your selection will be filled with the template formula

### Fill Blanks vs Fill Errors

- **Fill Blanks**: Fills only truly empty cells (no content at all)
- **Fill Errors**: Fills cells containing Excel error values (#REF!, #VALUE!, #DIV/0!, #N/A, etc.)

## Troubleshooting

### Add-in Not Appearing

1. Ensure the manifest file is named correctly (must end with `-manifest.xml`)
2. Verify the manifest is in the correct folder:
   ```bash
   ls ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
   ```
3. Completely quit and restart Excel
4. Try reinserting the add-in from Insert > My Add-ins

### Add-in Fails to Load

1. Check your internet connection (the add-in loads from GitHub Pages)
2. Verify GitHub Pages is accessible: visit https://mikedtcm22.github.io/excel_conditional_fill/
3. Clear the Office cache:
   ```bash
   rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/
   ```
4. Restart Excel

### Keyboard Shortcuts Not Working

1. Keyboard shortcuts may not be available in all Excel versions
2. Try using the ribbon buttons or context menu instead
3. Some shortcuts may conflict with other add-ins or system shortcuts

### Error Messages

**"No selection found"**: Select a range of cells before triggering the fill operation.

**"Active cell does not contain a formula"**: Position your cursor on a cell that contains the formula you want to copy before selecting your range.

**"No cells modified"**: All cells in your selection already have content (for Fill Blanks) or none have errors (for Fill Errors).

## Updating the Add-in

The add-in loads from GitHub Pages, so updates are automatic. Simply:

1. Refresh or restart Excel
2. The latest version will be loaded

If you encounter issues after an update, try clearing the Office cache as described in the troubleshooting section.

## Uninstalling

To remove the add-in:

1. Delete the manifest file:
   ```bash
   rm ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/FillGaps-manifest.xml
   ```
2. Restart Excel

## GitHub Pages URL

The add-in is hosted at:
- **Base URL**: https://mikedtcm22.github.io/excel_conditional_fill/
- **Manifest**: https://mikedtcm22.github.io/excel_conditional_fill/manifest-production.xml

## Support

For issues or feature requests, please visit:
https://github.com/mikedtcm22/excel_conditional_fill/issues
