# Sideloading FillGaps Add-in on Excel for Mac

This guide explains how to sideload the FillGaps add-in for testing and development on Excel for Mac.

## Prerequisites

1. Excel for Mac (Microsoft 365)
2. Development server running on localhost:3000
3. Trust self-signed certificate for HTTPS

## Step 1: Start the Development Server

```bash
npm start
```

This will:
- Build the add-in
- Start webpack-dev-server on https://localhost:3000
- Generate a self-signed certificate for HTTPS

## Step 2: Trust the Self-Signed Certificate

When you first access https://localhost:3000, you may see a security warning. To trust the certificate:

1. Open Safari or Chrome
2. Navigate to https://localhost:3000
3. Accept the certificate warning
4. Verify the page loads (you should see the taskpane.html or webpack-dev-server page)

## Step 3: Copy Manifest to Office Add-ins Folder

The Office Add-ins folder location on Mac:

```bash
~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
```

If the folder doesn't exist, create it:

```bash
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
```

Copy the manifest file:

```bash
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/FillGaps-manifest.xml
```

Or use this command from the project root:

```bash
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/FillGaps-manifest.xml
```

## Step 4: Open Excel and Insert the Add-in

1. Open Excel for Mac
2. Create a new workbook or open an existing one
3. Click on the **Insert** tab in the ribbon
4. Click **My Add-ins** dropdown
5. Under **Developer Add-ins**, you should see **FillGaps**
6. Click on **FillGaps** to load the add-in

## Step 5: Verify the Add-in Loads

After inserting the add-in:

1. Look for the **FillGaps** group in the **Home** ribbon tab
2. You should see a button labeled **Fill Gaps...**
3. Click the **Fill Gaps...** button
4. The task pane should open on the right side of Excel
5. Verify there are no console errors in the development tools

## Troubleshooting

### Add-in doesn't appear in My Add-ins

- Ensure the manifest file is in the correct location
- Restart Excel
- Check that the manifest.xml file is valid XML

### Certificate errors when clicking the button

- Ensure the development server is running
- Visit https://localhost:3000 in a browser and accept the certificate
- Clear Excel's add-in cache:
  ```bash
  rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/*
  ```
- Restart Excel

### Task pane shows blank or error

- Check the browser console in Excel's Developer Tools
- Verify webpack-dev-server is running
- Check that taskpane.html loads at https://localhost:3000/taskpane.html
- Review the manifest URLs to ensure they match the dev server

### Clearing the Add-in Cache

If you make changes to the manifest, you may need to clear Excel's cache:

```bash
# Clear Excel cache
rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/*

# Remove the manifest
rm ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/FillGaps-manifest.xml

# Copy updated manifest
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/FillGaps-manifest.xml
```

Then restart Excel.

## Useful Commands

```bash
# Quick sideload (copy manifest)
npm run sideload

# Start dev server
npm start

# Build production version
npm run build

# Run tests
npm test
```

## Development Workflow

1. Keep the dev server running with `npm start`
2. Make changes to source files
3. Webpack will automatically rebuild
4. Reload the task pane in Excel (close and reopen, or use Developer Tools to reload)
5. Test your changes

## Next Steps

Once the add-in is sideloaded and the ribbon button appears:
- Task Group 3 will implement the task pane UI
- Task Group 4 will add Office.js initialization and event handlers
- Task Groups 5-7 will implement the core fill logic
