# ORC Service Overview Macro

This repository contains the VBA macro code extracted from the Excel workbook
**"ORC_Service-Übersicht Makro.xlsm"**.

The workbook is used to manage firewall requests. The macro reads requests from
worksheets such as `AKH`, `WSK`, and `MAG`, checks them against lookup tables,
and generates new Excel files for any open or unprocessed requests.

The main workflow performed by the macro is:

1. Iterate over defined worksheets that hold firewall data.
2. For every request marked as **"Nein"** in the `Beantragt` column, export the
   data grouped by environment (`DEV`, `MIG`, `PROD`).
3. Create a new workbook for each environment containing only the unprocessed
   records. Separate files are generated for "Firewall" and "Server Firewall"
   requests.
4. Save the generated workbooks to an `Export` subfolder next to the original
   workbook.
5. If the file saves successfully, mark the original rows as processed and set
   the processing time and status.

If saving fails, the macro leaves the original data untouched so that the
request can be retried.

The VBA code for the macro lives in [`ThisWorkbook.vba`](ThisWorkbook.vba).
