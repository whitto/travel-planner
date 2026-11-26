# Travel Planner

Travel Planner is an iPhone-first progressive web app that imports an Excel itinerary and stores it on the device. It offers day and category filters, search, booking references, map links and offline access.

## Install on iPhone

Open the GitHub Pages site in Safari, tap **Share → Add to Home Screen**, then open Travel Planner and tap **Choose Excel file**. Pick your workbook from Files. Wait for **Offline ready** while connected before trying airplane mode. When **Update app** appears, tap it to reload the latest version. Updates preserve the saved itinerary.

## Workbook format

Use .xls or .xlsx with one itinerary item per row, headings in row 1 and dates in column A. A worksheet named **Daily Plan** takes priority. For that layout the app also accepts a Date header in the first 25 rows and dates like “3 Dec”. Use a four-digit starting year in the filename for dates without a year, and list those dates in calendar order so the app can handle a year rollover. Full dates can be Australian day/month/year, ISO year-month-day or Excel date cells. Remaining columns can have arbitrary headings. Settings lists any rows skipped because of missing or invalid dates.

Use **Settings → Refresh from Excel / Change spreadsheet** to open an updated workbook. Settings also exports and restores private JSON backups. Browsers or operating systems may clear local storage, so keep the original workbook.

Excel is parsed on the device with the bundled [SheetJS Community Edition 0.20.3](https://cdn.sheetjs.com/xlsx-0.20.3/xlsx-0.20.3.tgz). Apache-2.0 license notices are in `vendor/LICENSE`. There is no backend or analytics. Map and booking links open the selected external service and may need an internet connection.

## Development and deployment

Run `node --test tests/import.test.cjs` with Node.js. Synthetic .xls and .xlsx workbooks are in `examples/`.

Serve the repository with a static HTTP server on localhost for development. Deploy from `main` and `/(root)` through GitHub Pages for production. The manifest and service worker use relative URLs, including when the app lives under `/travel-planner/`. Increment the cache version in `sw.js` when releasing app-shell changes.
