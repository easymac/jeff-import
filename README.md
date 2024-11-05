# jeff import

A simple (throwaway) script to help parse the JN Excel sheets into Liftoscript.

## Getting started

1. Clone the repository
2. npm install
3. PREPARE THE SPREADSHEET!
4. node parse.js

## Preparing the sheet

A small amount of cleanup is required in Excel for two reasons:

1. The warm up rep ranges are formatted in Excel as DATES. This is by accident. So a rep range written as "2-3" is actually February 3rd, 2023 and stored in Excel as the number "44960". This is a problem for the importer.
2. The URLs need to be pulled out of the hyperlinks.

It's probably easiest just to open `Prepared_Sheet.xlsx` and see how it's done, but these are the changes:

1. Delete all the junk that isn't part of the training sheet. Actually, just make a new sheet & copy and paste the actual program into it.
2. Highlight all cells and unmerge them (Home > Merge & Center > Unmerge Cells) -- this will let you drag formulae in the next step
  1. Add a NEW COLUMN to the LEFT of the Warm-Up sets record.
  2. Add the following formula: `=IF(F4>10,TEXT(F4,"m-d"),IF(ISNUMBER(F4),F4,""))`
  3. Drag fill it down the column to the entire sheet.
3. Add a VB script to extract the URLs from hyperlinks. Here is the tutorial on how to do that: [https://www.excel-university.com/extract-url-from-hyperlink-with-an-excel-formula/](https://www.excel-university.com/extract-url-from-hyperlink-with-an-excel-formula/). If you are simply editing Prepared_Sheet, it should already be in there and you can skip this step.
4. Insert a column to the RIGHT of the Exercise column, and populate it with URLs using the following formula: `=IFERROR(URL(B4), "")`

Hopefully now your sheet matches the format in `Prepared_Sheet.xlsx`.

If you want, I can handle this step for you. Just ping me on Discord.

---

That should be it. I'm happy to run the script on your behalf. PRs, suggestions, bug reports, etc are welcome.

---

In general, this is not a good way to handle this process. Obviously. But it's easy and didn't require me to fuss about with sheetsjs.