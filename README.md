# Canvas Course Exporter

A static web page that reads a Canvas course export (`.imscc`) and lets the
user choose what to pull out of it:

| Section | What you get |
|---|---|
| **Assessments** → formatted Word documents | Classic Quizzes question banks · New Quizzes item banks · Quizzes with question groups · Quizzes with fixed questions (Classic and New) |
| **Course files** → original files, unchanged | Documents (Word, PDF, PowerPoint, Excel…) · Images · Audio, video and other files |
| **Course content** → Word documents | Pages · Assignments · Discussions and announcements · Rubrics · Course outline (modules, assignment groups, learning outcomes) |

Assessments are checked by default; everything else is opt-in. Everything
selected comes back as one `.zip` with a folder for each kind of content (or
the single file directly, if only one item is chosen).

All processing happens in the user's browser. The `.imscc` file is never
uploaded anywhere, so the site can be hosted on GitHub Pages with no server.

---

## Put it on GitHub Pages

1. Create a new repository on GitHub (for example `canvas-course-exporter`).
2. Upload the contents of this folder to the root of the repository
   (`index.html`, `css/`, `js/`, `vendor/`, `README.md`, `.nojekyll`).
   - On github.com: **Add file → Upload files**, drag everything in, commit.
     (The `.nojekyll` file is hidden on Mac/Linux; it's optional.)
   - Or with git:
     ```bash
     git init
     git add .
     git commit -m "Canvas Course Exporter"
     git branch -M main
     git remote add origin https://github.com/<you>/canvas-course-exporter.git
     git push -u origin main
     ```
3. In the repository go to **Settings → Pages**.
   Under **Build and deployment**, set **Source** to *Deploy from a branch*,
   choose **main** and **/ (root)**, then **Save**.
4. After a minute the site is live at
   `https://<you>.github.io/canvas-course-exporter/`.

### Run it locally

No build step. Open `index.html` directly in a browser, or serve the folder:

```bash
python3 -m http.server 8000
# then open http://localhost:8000
```

---

## Using it

1. In Canvas: **Settings → Export Course Content → Course → Create Export**,
   then download the `.imscc` file.
2. Drop the file on the upload area.
3. **What do you want to extract?** Tick the kinds of content you want. Each
   shows how many were found; anything not in the export is greyed out.
4. **Choose items.** Each ticked kind gets its own list (all items selected;
   use the filter box or All/None to narrow).
5. **Options**, then **Create download**.

### Options

Only the options for the ticked content are shown.

| Option | What it does |
|---|---|
| Answer key / Student copy / Both | Key marks correct answers in green with a ✓ and shows feedback; student copy leaves answer lines and response boxes instead. |
| One Word file per bank or quiz / All in one | Separate files go into a folder per assessment type. |
| Answer feedback | Per-answer, general, correct and incorrect feedback (answer key only). |
| Point values | Points per question and per question group. |
| Questions pulled from banks | For question groups that draw from a bank, list that bank's questions inside the quiz. |
| Answer summary table | Question number → correct answer table at the end of each bank or quiz. |
| Keep Canvas folder structure / One folder | How course files are laid out in the zip. |
| Include settings | Points, due/availability dates, submission type and assignment group at the top of assignment and discussion documents. |

### What the download looks like

```
<Course> - Canvas export.zip
├── Question banks (Classic)/
├── Item banks (New Quizzes)/
├── Quizzes with question groups/
├── Quizzes/
├── Course files/          originals, in Canvas's folder structure
├── Pages/
├── Assignments/           includes the attached rubric, if any
├── Discussions/
├── Announcements/
├── Rubrics/               landscape tables: criteria × rating levels
└── Course outline.docx
```

---

## Details and limitations

**Quiz question types:** multiple choice, true/false, multiple answers,
matching, fill in the blank, fill in multiple blanks, multiple dropdowns,
numerical (exact, margin and range), formula (formula, variables, tolerance,
example sets), essay, file upload, and text-only/passage items. Other New
Quizzes types (ordering, categorization, hot spot) are shown as far as the
export allows and flagged to check in Canvas.

**Formatting in Word:** bold, italic, underline, strikethrough,
super/subscript, lists, tables, links, and images from course files (PNG, JPG,
GIF, BMP embedded; SVG and WebP converted to PNG). Canvas equations appear as
their LaTeX text. Background colours and custom page styling are not kept.

**Not in a Canvas export, so can't be extracted:**
- student submissions, grades, quiz attempts, discussion replies
- New Quizzes item banks that no New Quiz in the course uses
- files linked from Google Drive, OneDrive, YouTube, Canvas Studio or other
  external tools (only files uploaded to Canvas Files are included)

**Large exports:** courses with a lot of video can produce exports of several
hundred MB or more. The browser handles them, but loading and zipping take
longer and need memory roughly the size of the export.

**Page numbers** in the Word footer are fields. Word fills them in
automatically; some other viewers only update them when printing.

---

## Project layout

```
index.html              page markup
css/styles.css          styles (light and dark mode)
js/parser.js            reads the .imscc; quiz and bank QTI → quiz model
js/content-parser.js    files, pages, assignments, discussions, rubrics,
                        modules, assignment groups, outcomes
js/html-to-docx.js      converts Canvas HTML into Word paragraphs/tables/images
js/word-builder.js      quiz/bank Word layout + shared document shell
js/content-builder.js   Word layout for pages, assignments, discussions,
                        rubrics and the course outline
js/extractors.js        list of extractable categories, image loading,
                        building the zip
js/app.js               page behaviour (upload, checklists, options, download)
vendor/                 JSZip 3.10.1 and docx 9.7.1 (MIT), bundled so the site
                        needs no outside CDN
```

### Adding another kind of content

Add an entry to `CQE.CATEGORIES` in `js/extractors.js` (`id`, `section`,
`label`, `folder`, and `items(pkg)` returning `{id, title, meta, ref}`), then
add a job for that `id` in `buildJobs()` in the same file. The checklist,
item list and counts in the page pick it up automatically.

---

## Data handling

Nothing is transmitted or stored: the file is read with JavaScript in the
browser tab and the output is created in the same tab. Exam content and
course materials are still sensitive, so handle the downloads with the same
care as the original Canvas content and follow your institution's data
policies.

## Licenses

Bundled libraries: [JSZip](https://stuk.github.io/jszip/) (MIT or GPLv3, used
here under MIT) and [docx](https://docx.js.org/) (MIT). Their license files
are in `vendor/`.
