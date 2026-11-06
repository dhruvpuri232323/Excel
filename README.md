# Sheetah: Excel, for a purpose

Sheetah is a self-paced Excel course that runs as a website. You can host it free on GitHub Pages. It needs no server, no payments and no database.

- **Two courses.** *C1 · Excel Foundations* (8 modules) and *C2 · Formulas & Functions* (9 modules, released in parts: Modules 1–3 are out now, the rest are marked "coming soon").
- **Free for anyone.** Modules 1–2 of *Excel Foundations*, Module 1 of *Formulas & Functions*, plus every study tool, are open to all visitors. The tools are the function library, shortcut trainer, playground, weekly challenge, interview prep, daily review, mistakes log, study plan, journal, insights and practice tests.
- **Students with a login go further.** Modules 3–8 open one at a time as the student passes each module test, then the course project, the final exam and the certificate. A student who already knows a module can **test out** of it.
- **You give out the logins.** Each student gets an email and password from you, and nobody pays inside the app.

---

## 0. Updating a site that's already live

If your site is already on GitHub Pages, you only upload the new files. **Your student list stays as it is**: the update has no `data` folder, so `data/students.json` isn't touched.

1. Unzip the update. Open the unzipped folder: you'll see `index.html`, `css`, `js`, `content`, `img`, `vendor` and a few other files.
2. Upload in **three drags**, because GitHub's web uploader takes up to 100 files at a time:
   1. In your repository, click **Add file › Upload files**. Drag in everything **except** `content` and click **Commit changes**.
   2. In your repository, click the `content` folder to open it, then **Add file › Upload files**. From your unzipped `content` folder, drag in the `c2` folder, `catalog.json` and `library.json`, and click **Commit changes**.
   3. Still inside `content`, click **Add file › Upload files** again, drag in the `c1` folder and click **Commit changes**.
   - Files with the same name are replaced. That's what you want.
3. Wait a minute, then open your site and press **Ctrl+Shift+R** (Cmd+Shift+R on a Mac) once, so the browser fetches the new version.
4. Sign in as admin (sign out and in again if you were already signed in). Your admin login picks up any new course by itself; nothing else to do for you. See **Giving students a course** below for your students.
5. The **Admin** page opens with a five-step setup checklist. Do the steps there: add your WhatsApp number or email under **Site details** (every "Ask your teacher" button uses it), then **Publish**.

Students keep all their progress. Anyone who already used Sheetah skips the new setup questions; they can run them from **Settings › Run setup again**.

### Giving students a course (for example, Formulas & Functions)

Each course has its own lock, and a student's login holds the keys to the courses you chose for them.

- **New students:** in **Admin › Students › Add a student**, tick every course they should get. When you paste a whole class, tick the courses under **Courses for this class**.
- **Students you already have:** in the student list, click **Change** next to their courses, tick the new course, and click **Save and make a new password**. A login can't be changed without a new password, so send them the new one from the next screen (WhatsApp, Email or Copy). Their progress stays exactly as it was.
- Then **Publish**. Until you publish, students keep their old access.
- Students without a course still see its free modules, and a note telling them to ask you for it.

## 1. Go live on GitHub Pages for the first time (about 10 minutes)

1. Sign in at **github.com** and click **New repository**.
   - Name it, for example, `sheetah`.
   - Make it **Public**. GitHub Pages is free for public repositories. The locked lessons stay encrypted, so being public is fine.
   - Click **Create repository**.
2. Put the files in the repository, in **three drags** (GitHub's web uploader takes up to 100 files at a time):
   1. On the new repository page, click **uploading an existing file**. Drag in everything inside this folder **except** `content`: that is `index.html`, `css/`, `js/`, `data/`, `img/`, `vendor/`, `sw.js`, `manifest.webmanifest`, `.nojekyll` and `README.md`. Click **Commit changes**.
   2. Click **Add file › Create new file**, type `content/x` as the name, and commit; this makes the `content` folder. Open `content`, click **Add file › Upload files**, drag in the `c2` folder, `catalog.json` and `library.json`, and commit. Then delete the file `x`.
   3. Still inside `content`, click **Add file › Upload files**, drag in the `c1` folder and commit.
   - Drag the *contents* of this folder, not the folder itself, so that `index.html` sits at the top of the repository.
   - Can't see `.nojekyll`? Files starting with a dot are hidden on Mac and Windows. Press Cmd+Shift+. on a Mac, or turn on View › Hidden items on Windows. The site still works without it.
3. Go to **Settings › Pages**. Under *Build and deployment*, choose **Deploy from a branch**, pick branch **main** and folder **/ (root)**, then click **Save**.
4. Wait a minute and refresh the Pages screen. It shows your address, for example `https://yourname.github.io/sheetah/`. That is the link you share.

## 2. The Admin page

Sign in with the admin email and password, then open **Admin** in the side panel. It has five tabs and a setup checklist that ticks itself as you go:

1. **Change your admin password** (Students tab, bottom). The temporary one stops working once you publish.
2. **Add your WhatsApp number or email** (Site details tab).
3. **Add your first student** (Students tab).
4. **Remove the demo login**: click **Remove** next to *Demo Student* once you've tried the course as a student.
5. **Publish** (see section 3).

What each tab does:

- **Students.** Add one student or paste a whole class (`Name, email, WhatsApp number`, one per line). After adding, Sheetah shows each login with **WhatsApp** and **Email** buttons that send the site link, the email, the password and a two-line "how to start". Passwords are never stored where you can read them again, so send them then. **New password**, **Pause** and **Remove** work per student. The **All modules open** tick opens every module for one student (for someone who already knows Excel); a switch opens everything for everyone.
- **Class progress.** Students send you a progress file from *Settings › Your data › Progress file*. Drop the files here for one table: lessons done, tests passed, last studied and the topics they need to practise. Nothing is uploaded anywhere.
- **Site details.** Your contact details, and an **announcement** shown on every student's home page (for example "No class on Diwali").
- **Certificates.** When a student passes, their certificate page has a button that sends you the certificate ID. Paste it here and publish: the public verify page then shows the certificate as genuine.
- **View as a student.** Preview the app as a visitor, a new student or a student halfway through, without signing out. Nothing in a preview is saved.

## 3. Publishing

The student list and site details live in `data/students.json`. Changes on the Admin page are a draft until you publish; a yellow banner says how many changes are waiting.

**Option A: upload one file (no setup).** Click **Publish now** (or **Download students.json**) and follow the three steps on screen: download the file, open the `data` folder in your repository, **Add file › Upload files**, then **Commit changes**. Wait about a minute and click **I've uploaded it**.

**Option B: one-click publishing (set up once).** Under **One-click publishing** on the Students tab, click **How do I get a token?** and follow the steps. With a token saved, **Publish now** updates the file for you. The token is saved only in that browser on that computer; revoke it on GitHub if the device is lost.

## 4. What a student experiences

- **Setup first.** One question per screen: light or dark, text size, name, why they're learning, where they're starting, Sheetah's tone, study days and time, Windows or Mac. It ends with a finish date and a choice: take the 7-stop tour or start lesson 1. Students who already know some Excel are offered a 10-minute placement check.
- **The learning map.** Every module is a stop, with "You are here", and the diamonds between stops are module tests. Hovering a stop lists its lessons and the time left. Each module has its own page: lessons, test, topics that jump to the exact slide, and downloads.
- **Study plan:** pick a start date and a finish date (plus study days and any days off) and the day-by-day plan, calendar and "Today" list build themselves, re-spreading automatically after a missed day.
- **Lessons** have five parts, each with its own web address so Back works: **Lecture** (visual slides: one idea, a diagram, short points, a real-life example and words to know; or read it all as one page of notes), **Demo** (Sheetah does it on a live sheet), **Lab** (tasks on a working sheet, three hints, "What Sheetah checked", and the solution after three tries), **Quiz** and **Recap**. ☰ Outline lists every slide and task; ⤢ Full screen hides everything else.
- **More than one course.** Home shows a switch (C1 Foundations · C2 Formulas) when more than one course is live; the study plan, Today list and notes follow the course chosen there, or the one whose lesson was opened last.
- **Help everywhere.** Small **?** buttons, underlined Excel words with their meanings, a searchable Help centre, a one-page "How Sheetah works" guide with a one-minute video, search across everything (**Ctrl+K**), and **Ask your teacher** buttons that open WhatsApp or email with a ready message.
- **Tests.** Module tests take about 25 minutes, pass mark 70%. A failed test lists the lessons to revisit; the retake opens after 24 hours, or straight away once those lessons are redone. The course project and the final exam then lead to a PDF certificate with an "Add to LinkedIn" link and a verify page.
- **Practice and habits.** Practise any lesson with fresh numbers, a weekly real-world challenge, 20 interview questions with a timed mock, daily review cards, the mistakes log, a study plan with a calendar file, a journal, insights, notes (downloadable as Word) and saved items.
- **Settings** apply instantly: theme, accent colour, text size, line spacing, an easy-reading font, high contrast, reduced motion, celebrations, side panel, spacing, mascot, tone, goals, reminders, Windows or Mac keys, avatar, and menus in Hindi.

## 5. Good to know (limits of a server-free site)

- **Progress is saved in each student's browser.** If a student switches laptops or clears browser data, their progress doesn't follow them. *Settings › Your data* has a backup, a restore and a step-by-step "move to another laptop" helper.
- **How the lock works.** Locked lessons and files are encrypted. A student's password unlocks a copy of the course key, so sharing a password shares access, and the module-by-module order is a guide that a determined student could get around. Pause or remove a login if needed.
- **GitHub Pages terms.** Pages is meant for non-commercial sites. That fits, because nobody pays through the site. If you later sell access, move to a host such as Netlify or Cloudflare Pages; the files stay the same.
- **Screenshots are illustrations.** They are drawn to look like Microsoft Excel for Windows (Microsoft 365) and are marked "Illustration · Sheetah". Menus can look slightly different on a Mac or in older versions.
- Sheetah is not affiliated with Microsoft. Excel is a trademark of Microsoft Corporation.

## 6. Updating the course later

The lesson text, the encryption keys and the build scripts live in a separate **private** `source` folder. **Never upload `source/`, and especially never upload `source/secrets/`, to the public repository:** `keys.json` would let anyone decrypt the course. `source/BUILD.md` explains how to rebuild the site after editing lessons.
