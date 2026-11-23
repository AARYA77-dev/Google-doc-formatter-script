# Format QA Issue Report

Google Apps Script for formatting a QA issue report in Google Docs. It is designed to run as a **container-bound script** in the Google Doc it formats.

## What it does

- Adds a title banner and a header and footer.
- Applies a consistent font, page margins, and page background.
- Styles issue list items and adds optional separator lines.
- Resizes oversized screenshots and styles screenshot tables as cards.
- Converts plain-text URLs into clickable links.
- Adds a status dropdown beneath each issue title. The available statuses are **Open**, **Hold**, **Not fixed**, **Not a bug**, and **Fixed**.
- Optionally adds issue headings to the Google Docs outline for quick navigation.
- Offers a temporary status filter in a document tab of the same Google Doc, without changing the original report.

The script formats the document; it does not rewrite issue wording.

## Before you start

1. Make a copy of the Google Doc (**File → Make a copy**) and run the script on that copy first.
2. Make sure each issue heading is a real Google Docs list item, not a number typed as ordinary text. The script recognizes list items as issues.
3. Status dropdowns, document tabs, and issue filtering use the **Google Docs API**. If Google reports that the API is disabled, open the linked project in Google Cloud Console and enable **Google Docs API** under **APIs & Services → Library**.

## Install the script

1. Open the Google Doc to format. If the report is a `.docx`, upload it to Drive and choose **Open with → Google Docs** first.
2. In the Doc, select **Extensions → Apps Script**. This opens a script project bound to this Doc.
3. Replace the editor contents with the contents of [`FormatQAReport.gs`](./FormatQAReport.gs), then save the project.
4. In the Apps Script editor, select `formatQAReport` from the function list and click **Run**.
5. Review and allow the requested Google permissions. The script needs access to the open document and Google Docs API requests. Filtering no longer uses Drive access.
6. Return to the Doc. If the options dialog did not appear, make sure the Doc is open and reload it. The **Format report** menu should appear after reloading.

The Apps Script editor may report **Execution completed** while the options dialog is open in the Google Doc. Switch back to the Doc to continue.

## Format the report

Use **Format report → Format QA report…** in the Doc. Choose the options and select **Format document**.

| Option | What it does | Default |
| --- | --- | --- |
| List style | Choose numbers, bullets, letters, or Roman numerals for issue list items. | Numbers |
| Document outline | Add the report title and issue headings to the outline sidebar without creating per-issue document tabs. | On |
| Separator line | Add a horizontal line before each issue after the first. | On |
| Image size | Keep images unchanged or shrink images wider than the page. Images are never enlarged. | Keep unchanged |
| Page background | Apply the selected background color or keep the page white. | Background on |
| Banner colour | Choose a preset or use the color picker. | Dark blue |

Status dropdowns are added automatically when formatting. To add or refresh them without running the other formatting steps, use **Format report → Add issue status dropdowns**.

### Temporary status filter

Google Docs does not provide a supported way for Apps Script or the Docs API to hide arbitrary paragraphs in a document. **Format report → Filter Issues** therefore builds a temporary filtered tab inside the same Google Doc. The original tab is never filtered or edited.

- Choose **Open**, **Hold**, **Fixed**, **Not a Bug**, or **Not fixed**. A progress dialog reports the number of issues copied; when complete, select **Open filtered view in this document** to switch to the new tab in the current browser window.
- Choose another status to refresh the same temporary filter tab. The issue sections are copied in batches, including their formatting, images, and tables. Status labels in this temporary tab are plain text; the original dropdowns remain unchanged in the original tab.
- **Show All Issues** links back to the unchanged original tab.
- **Restore Original View** removes the temporary filtered tab and links back to the original tab. It does not remove any of the report's original issue tabs.
- A filter requires a readable status dropdown for every issue. If one is missing, run **Format report → Add issue status dropdowns** and retry.

The filter never deletes or edits original issue content. Temporary Drive copies created by an earlier version are not automatically removed; remove those old copies manually if they are no longer needed.

### Outline navigation

When enabled, the report title appears as a top-level outline item and each issue appears beneath it with its selected list-style marker (for example, `1.`). Status dropdowns are kept in normal-text paragraphs and are not included in the outline. Use **View → Show outline** in Google Docs to open the navigation sidebar and select an issue. The issue marker is written into the heading text and its list bullet is removed to prevent duplicate numbering. Existing issue content and formatting are retained; the report stays in one main document tab and no per-issue tabs are created. Existing tabs from an earlier script version are left untouched.

## Configuration

The `CFG` object near the top of [`FormatQAReport.gs`](./FormatQAReport.gs) contains settings for the font, page size and margins, colors, header and footer text, status dropdown choices, and image sizing. Edit these values in Apps Script and save before running the formatter.

The default page size is US Letter. The source includes a comment with the A4 dimensions if you need to change it.

## Troubleshooting

- **The Format report menu is missing:** Save the script, return to the Doc, and reload the page.
- **The Filter Issues submenu is missing:** Confirm the updated `onOpen` function is saved in this Doc's bound Apps Script project, then reload the Doc. `onOpen` runs when the Doc is opened; do not run it from the script editor.
- **The dialog does not open:** Keep the target Doc open in a browser tab, reload it, and try **Format report → Format QA report…** from the Doc.
- **No issue list items were found:** Use Google Docs' list formatting on issue headings; typed prefixes such as `1.` are not list items.
- **A step failed:** The dialog or return message names the failed step and includes its error. For details, open **View → Logs** in the Apps Script editor.
- **`Could not open the requested document tab`:** Save the latest script in the bound project, reload the Doc, and retry. If it persists, note whether a `Filtered Issues` tab appeared and include the tab ID shown in the error or logs when reporting it.
- **A Google Docs API request reports that the API is disabled or access is denied:** Enable the Google Docs API for the Apps Script project's linked Cloud project, then authorize the script again. Your Google Workspace administrator may restrict access to Apps Script or the Docs API.
- **`You do not have permission to call DriveApp.getFileById`:** This version of the filter does not use Drive. Save the updated script in the bound project and reload the Doc. If the error remains, check that the Doc is using the updated script and that no other code in the project calls `DriveApp`.
- **Formatting is not as expected:** Revert to the copy you made before running the script, adjust the settings in `CFG` or the dialog, and try again. When separator lines are enabled, running formatting repeatedly can add additional lines.

## Running without the dialog

`formatWithDefaults` applies the script's default formatting choices without opening the options dialog. Outline navigation is on by default. Run it from the Apps Script editor only when you want those defaults.
