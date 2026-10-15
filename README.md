# Project Viewer (static HTML)

One HTML file that opens Microsoft Project schedules in the browser: a task table with the outline from Project, a Gantt timeline with dependency arrows and critical tasks, task details, search and filters, and CSV export.

**There is no server.** The file is read by JavaScript in the visitor's browser and is never uploaded, so you can host `index.html` anywhere that serves static files.

## Supported files

- `.mpp` / `.mpt` saved by **Project 2010, 2013, 2016, 2019, 2021 or Microsoft 365**
- **Project XML** (`.xml`) from any version: in Project, use *File › Save As* and choose *XML Format*

Files saved in the Project 2007-or-earlier format show a message asking the user to re-save them. Password-protected files must have the password removed first.

## Put it online

Upload `index.html` on its own; nothing else is needed. Some options:

- **Netlify Drop:** drag the folder onto https://app.netlify.com/drop and you get a URL in seconds.
- **Cloudflare Pages** or **GitHub Pages:** add `index.html` to a repository or project and publish it.
- **Azure Static Web Apps** or **SharePoint/IIS/any web host:** copy `index.html` into the site.
- **No hosting:** double-click `index.html` and it works offline in Chrome, Edge, Firefox or Safari.

The page loads the IBM Plex Sans font from Google Fonts. If it can't reach Google (offline or blocked), it falls back to the system font and still works.

## Accuracy

The `.mpp` reader is a JavaScript port of the relevant parts of MPXJ, the open-source library behind most non-Microsoft Project readers. It was checked against MPXJ on 164 sample `.mpp` files from Project 2010 to Microsoft 365. Task names, outline, WBS, dates, durations, % complete, critical flags, slack, baselines, dependencies, notes and resource names matched on 163 of them; the remaining file differs only in how one inserted-subproject row is styled.

## Editing the source

`index.html` is built from the files in `source/`:

```
source/viewer-template.html   the page (layout, styles, Gantt)
source/mpp-reader.js          the .mpp / XML reader
source/tables.js              field tables generated from MPXJ
source/build-html.mjs         combines them:  cd source && node build-html.mjs  ->  dist/index.html
```

## License

The reader code (`source/mpp-reader.js`, `source/tables.js` and the reader section of `index.html`) is derived from MPXJ (Copyright Jon Iles) and is licensed under the GNU Lesser General Public License 2.1; see `LICENSE-LGPL-2.1.txt`. In practice this lets you use and host it freely, including commercially, as long as the reader's source stays available (it is plain, unminified JavaScript inside `index.html`) and you keep this notice.
