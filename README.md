# Sheetah: Excel, for a purpose

Sheetah is a self-paced Excel course that runs as a website. You can host it free on GitHub Pages. It needs no server, no payments and no database.

- **Free for anyone.** Modules 1–2 of *C1 · Excel Foundations*, plus every study tool, are open to all visitors. The study tools are the function library, the shortcut trainer, the playground, daily review, the error log, the study planner, the journal, insights and practice tests.
- **Students with a login get everything.** That means Modules 3–8, module tests, the course project, the final exam and the certificate.
- **You give out the logins.** Each student gets an email and password from you, and nobody pays inside the app.

---

## 1. Go live on GitHub Pages (about 10 minutes)

1. Sign in at **github.com** and click **New repository**.
   - Name it, for example, `sheetah`.
   - Make it **Public**. GitHub Pages is free for public repositories. The locked lessons stay encrypted, so being public is fine.
   - Click **Create repository**.
2. Put the files in the repository. GitHub's web uploader takes up to 100 files at a time, and Sheetah has about 120, so upload in **two drags**:
   1. On the new repository page, click **uploading an existing file**. Drag in everything inside this folder **except** `content`: that is `index.html`, `css/`, `js/`, `data/`, `img/`, `vendor/`, `sw.js`, `manifest.webmanifest`, `.nojekyll` and `README.md`. Click **Commit changes**.
   2. Click **Add file › Upload files**, drag in the `content` folder and click **Commit changes**.
   - Drag the *contents* of this folder, not the folder itself, so that `index.html` sits at the top of the repository.
   - Can't see `.nojekyll`? Files starting with a dot are hidden on Mac and Windows. Press Cmd+Shift+. on a Mac, or turn on View › Hidden items on Windows. The site still works without it.
   - Prefer an app? GitHub Desktop (desktop.github.com) works too: clone the repository, copy the files in, **Commit to main**, then **Push origin**.
3. Go to **Settings › Pages**.
   - Under *Build and deployment*, choose **Deploy from a branch**.
   - Pick branch **main** and folder **/ (root)**, then click **Save**.
4. Wait a minute and refresh the Pages screen. It shows your address, for example `https://yourname.github.io/sheetah/`. That is the link you share.

Want a custom domain such as `learn.yourname.in`? Add it under **Settings › Pages › Custom domain**, then follow GitHub's DNS steps.

## 2. First thing after going live: secure the admin login

1. Open your site and click **I have a student login**. Sign in with the admin email and the **temporary admin password** you were given.
2. Go to **Admin › Your admin password**, type a new password and click **Change admin password**.
3. **Publish** the student list (see section 4). The temporary password stops working once the new list is live.
4. Remove the **Demo Student** login: in the student table, click **Remove**, then publish. Do this once you've finished trying the course as a student.

## 3. Add students

On the **Admin** page:

- **One student:** type their name and email. Click **New** to generate a password, or type your own. Tick the course and, if you like, set an **Access until** date. Click **Add student**.
- **A whole class:** paste one student per line as `Name, email` (a password is optional) and click **Add everyone**.
- A box then shows each login ready to copy or download as CSV. **Copy it now.** Passwords are never stored where you can read them again. If a student forgets theirs, click **New password**.
- **Pause** stops a student signing in. **Remove** deletes the login. Their progress stays on their own device.

### Logins only work after you publish

The student list lives in `data/students.json`. Each change on the Admin page has to be published before it takes effect. There are two ways to do this.

## 4. Publish the student list

**Option A: download and upload (no setup)**

1. On the Admin page, click **Download students.json**.
2. In your repository, open the `data` folder and click **Add file › Upload files**.
3. Drop the downloaded `students.json` in and click **Commit changes**.
4. Wait about a minute for GitHub Pages to update. The new logins then work.

**Option B: one-click publishing (set up once)**

1. On GitHub, open **Settings › Developer settings › Personal access tokens › Fine-grained tokens** and click **Generate new token**.
   - For **Repository access**, choose *Only select repositories* and pick your Sheetah repository.
   - Under **Permissions**, set **Contents** to *Read and write*.
   - Generate the token and copy it.
2. On the Admin page, open **One-click publishing**. Enter your GitHub username, the repository name, the branch (`main`) and the token.
3. Click **Save GitHub details**. From now on, **Publish to GitHub** updates the student list directly.

The token is saved only in that browser on that computer. Use the token on your own device only, and revoke it on GitHub if that device is lost.

## 5. What each student experiences

- They open your link. The first time, they give a first name, choose study days and a time, pick a daily goal and pick a "tone": warm, cheeky, strict coach or Hinglish.
- Each lesson has five parts: **Lecture**, **Demo**, **Lab**, **Quiz** and **Recap**.
  - **Lecture:** short slides with Excel screenshots and a check question.
  - **Demo:** Sheetah performs the steps on a live grid.
  - **Lab:** hands-on tasks on a working spreadsheet in the browser, with three hints per task. There is also an optional **Real Excel task**: the student downloads a workbook, does the task in real Excel, and uploads it to be checked automatically.
  - **Quiz:** five questions.
  - **Recap:** review cards go into a spaced-repetition deck.
- Every lesson has downloads: Word notes with screenshots, a practice workbook and a solution workbook.
- **Module tests:** 25 minutes, pass mark 70%, retake after 24 hours.
- **Course project:** an uploaded workbook, checked against a rubric. Some formula checks swap in hidden data sets. The student gets two submissions.
- **Final exam:** 90 minutes, 40 questions and 6 practical tasks, pass mark 75%, up to three attempts 48 hours apart.
- **Grade:** 60% exam and 40% project, giving a printable certificate. A grade of 90% or more earns a distinction.
- **Habit tools** (inspired by CAT prep apps):
  - A daily checklist.
  - A streak with free "freeze" days.
  - Reasons logged for missed days.
  - An **error log** that brings mistakes back after 1, 3 and 7 days.
  - A journal check-in.
  - A weekly review.
  - Practice "mock" tests with target scores and a score sheet.
  - An insights dashboard: heatmap, trends and mastery map.
  - Reminders.

## 6. Good to know (limits of a server-free site)

- **Progress is saved in each student's browser.** If a student switches devices or clears browser data, their progress doesn't follow them. Ask students to use **Settings › Back up your progress** now and then: **Download backup** saves a file, and **Restore a backup** loads it on another device.
- **Certificates can't be checked online.** Each certificate carries an ID and the student's results, but there is no server to look IDs up. If certificates matter for hiring, keep your own list of who passed.
- **How the lock works.** Locked lessons and files are encrypted. A student's password unlocks a copy of the course key, so sharing a password shares access. The same goes for a student copying the key out of their browser. This is a determined-cheater gap that every server-free site has. Pause or remove a login if needed.
- **GitHub Pages terms.** Pages is meant for non-commercial sites. That fits here, because nobody pays through the site. If you later sell access, move to a host such as Netlify or Cloudflare Pages; the files stay the same.
- **Screenshots are illustrations.** They are drawn to look like Microsoft Excel for Windows (Microsoft 365) and are marked "Illustration · Sheetah". Menus can look slightly different on a Mac or in older versions.
- Sheetah is not affiliated with Microsoft. Excel is a trademark of Microsoft Corporation.

## 7. Updating the course later

The lesson text, the encryption keys and the build scripts live in a separate **private** `source` folder. **Never upload `source/`, and especially never upload `source/secrets/`, to the public repository:** `keys.json` would let anyone decrypt the course. `source/BUILD.md` explains how to rebuild the site after editing lessons.
