# Task Group 2 Verification Checklist

Use this checklist to verify that Task Group 2 has been implemented correctly.

## Automated Verification

### 1. Build Verification
```bash
npm run build
```

**Expected Result:**
- Build completes successfully in ~1 second
- No TypeScript errors
- Output shows webpack compiled successfully
- Files created in `dist/` directory

### 2. Test Verification
```bash
npm run test:manifest
```

**Expected Result:**
- All 6 tests pass
- Test suite completes in < 1 second
- No test failures or errors

### 3. Full Validation
```bash
npm run validate
```

**Expected Result:**
- Build completes successfully
- All 6 manifest tests pass
- No errors in either step

---

## Manual Verification (Excel for Mac)

### 4. Sideload the Add-in
```bash
npm run sideload
```

**Expected Result:**
- Message confirms manifest copied
- File exists at: `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/FillGaps-manifest.xml`

### 5. Start Development Server
```bash
npm start
```

**Expected Result:**
- Webpack dev server starts on https://localhost:3000
- Browser opens automatically
- Certificate warning may appear (expected for self-signed cert)

### 6. Trust Self-Signed Certificate

In browser (Safari or Chrome):
1. Navigate to https://localhost:3000
2. Accept/trust the certificate warning
3. Verify page loads (may show webpack dev server page)

**Expected Result:**
- Browser shows page from localhost:3000
- No persistent certificate errors

### 7. Open Excel and Insert Add-in

In Excel for Mac:
1. Close Excel completely (if open)
2. Open Excel
3. Create new blank workbook
4. Click **Insert** tab in ribbon
5. Click **My Add-ins** dropdown
6. Look for **Developer Add-ins** section
7. Click **FillGaps** to insert

**Expected Result:**
- FillGaps appears under Developer Add-ins
- Clicking FillGaps loads the add-in
- No error messages

### 8. Verify Ribbon Button

After inserting the add-in:
1. Look at the **Home** ribbon tab
2. Scroll to find the **FillGaps** group
3. Verify button labeled **Fill Gaps...** exists
4. Hover over button to see tooltip

**Expected Result:**
- FillGaps group appears in Home ribbon
- Fill Gaps... button is visible
- Tooltip shows: "Open FillGaps settings to fill formulas into blank and error cells only"

### 9. Test Task Pane Opening

With development server running:
1. Click the **Fill Gaps...** button in ribbon
2. Task pane should attempt to open on right side of Excel
3. Check for any error messages

**Expected Result:**
- Task pane opens on right side
- May show basic placeholder content (from taskpane.html)
- No error dialogs
- No certificate errors (if cert was trusted in step 6)

### 10. Verify No Console Errors

If Excel supports Developer Tools:
1. Open Developer menu (if available)
2. Check browser console for errors
3. Look for manifest-related errors

**Expected Result:**
- No manifest validation errors
- No HTTPS/certificate errors
- Office.js loads successfully

---

## Verification Checklist

Mark each item as you verify it:

### Automated Tests
- [ ] `npm run build` completes successfully
- [ ] `npm run test:manifest` all 6 tests pass
- [ ] `npm run validate` completes without errors

### Sideloading
- [ ] `npm run sideload` copies manifest successfully
- [ ] Manifest file exists at wef folder location

### Development Server
- [ ] `npm start` starts dev server on https://localhost:3000
- [ ] Browser can access localhost:3000 with trusted cert

### Excel Integration
- [ ] FillGaps appears in Insert > My Add-ins > Developer Add-ins
- [ ] Add-in can be inserted into Excel workbook
- [ ] FillGaps group appears in Home ribbon
- [ ] Fill Gaps... button is visible in ribbon
- [ ] Button tooltip displays correctly

### Task Pane
- [ ] Clicking Fill Gaps... button opens task pane
- [ ] Task pane loads content from localhost:3000/taskpane.html
- [ ] No manifest errors in console
- [ ] No HTTPS certificate errors

---

## Troubleshooting

### Issue: Add-in doesn't appear in My Add-ins

**Solutions:**
1. Verify manifest is in correct location:
   ```bash
   ls -la ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
   ```
2. Restart Excel completely
3. Check manifest XML is valid (should pass validation tests)

### Issue: Certificate errors when clicking button

**Solutions:**
1. Ensure dev server is running (`npm start`)
2. Visit https://localhost:3000 in Safari and accept certificate
3. Clear Excel cache:
   ```bash
   rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/*
   ```
4. Restart Excel

### Issue: Task pane shows blank or error

**Solutions:**
1. Check webpack-dev-server is running
2. Verify https://localhost:3000/taskpane.html loads in browser
3. Check manifest URLs match dev server
4. Look for errors in browser console

### Issue: Ribbon button doesn't appear

**Solutions:**
1. Verify VersionOverrides section exists in manifest
2. Check manifest.xml has ribbon group configuration
3. Restart Excel
4. Try removing and re-inserting the add-in

---

## Success Criteria

Task Group 2 is considered successfully verified when:

1. ✓ All 6 manifest tests pass
2. ✓ Manifest validates against Office.js schema
3. ✓ Ribbon button appears in Excel for Mac
4. ✓ Clicking button opens task pane (even if blank)
5. ✓ No console errors related to manifest loading

---

## Next Steps

Once all items in this checklist are verified:
- Task Group 2 is complete
- Ready to proceed to Task Group 3: Task Pane HTML/CSS Structure
- The manifest and ribbon integration layer is ready for UI development

---

## Notes

- The task pane content will be minimal/placeholder until Task Group 3 implements the UI
- Icons are currently simple placeholders and will be replaced with actual branding later
- Development server must remain running while testing the add-in
- Self-signed certificate warning is normal for local development
