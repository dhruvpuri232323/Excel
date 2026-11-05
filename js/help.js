// Help that sits where it's needed: "?" tips, glossary words in lessons, the help centre's articles,
// and the "Ask your teacher" message.
import { store } from './store.js';
import { session } from './content.js';
import { site, teacher } from './site.js';
import { esc, md, $, $$, modal, toast, popover, waLink, mailLink } from './ui.js';

export const HELP_TIPS = {
  override: 'Ticked: every module is open for this student, without passing the tests in order. Useful for someone who already knows Excel. It goes live when you publish.',
  parts: 'Each lesson has five parts in order: **Lecture**, **Demo**, **Lab**, **Quiz** and **Recap**. A part opens when you finish the one before it.',
  segment: 'A lecture is split into short segments of a few slides. Each segment ends with one quick check question.',
  streak: 'Days in a row you studied. A finished lesson, 5 reviews or 10 minutes of study counts. Two free streak freezes a month cover a missed day.',
  minutes: 'Minutes you actively studied today, against your daily goal. Change the goal in Settings › Learning.',
  due: 'Review cards and past mistakes that are due today. About ten minutes of review keeps them fresh.',
  map: 'Each stop is a module. The diamonds between stops are module tests: pass one (70% or more) and the next module opens.',
  testrules: 'About 25 minutes: 15 questions and 2 practical tasks on a live sheet. Pass mark 70%. If you don’t pass, the retake opens after 24 hours, or straight away once you redo the lessons it suggests.',
  testout: 'Already know this module? Pass its test and Sheetah marks it **tested out** and opens the next module. Its lessons stay open for review.',
  mastery: 'Each lesson moves from New to Learning, Familiar, Proficient and Mastered as you review its cards. It dims again when cards are due.',
  cards: 'Cards come back just before you’d forget them. Rate honestly: **Again** brings a card back soon, **Easy** pushes it far out.',
  real: 'Optional but useful: do the same task in real Excel, upload the file, and Sheetah checks it at once.',
  inputs: 'Sheetah also tries your formula with different numbers, so it only passes if it points at the right cells.',
  placement: 'A 10-minute mix of questions from every module. It suggests where to start and which modules you could test out of. It doesn’t unlock anything by itself.',
  project: 'A real workbook you build in Excel and upload. Pass mark 70%, two submissions, and it counts for 40% of your course grade.',
  exam: '90 minutes, 40 questions and 6 practical tasks. Pass mark 75%, up to 3 attempts 48 hours apart. It counts for 60% of the course grade.',
  saved: 'Progress is saved in this browser on this laptop. Use Settings › Your data to back it up or move it to another laptop.',
  outline: 'Every slide, task and question in this lesson. Ticks show what you’ve done; click one to jump there.',
};

const P = p => `img/guide/${p}.png`;
export const ARTICLES = [
  { id: 'start', title: 'How a lesson works', tags: 'lecture demo lab quiz recap parts', img: P('lesson'), body: ['Every lesson takes about 30 minutes and has five parts, always in this order:', '**1. Lecture** (about 9 minutes): short slides with pictures of Excel. Each segment ends with one quick check.', '**2. Demo** (about 4 minutes): press Play and watch Sheetah do it on a live sheet, or step through it.', '**3. Lab** (about 10 minutes): your turn. Do each task on the practice sheet and press **Check**. Hints are there if you need them.', '**4. Quiz** (about 4 minutes): five questions. Wrong answers go to your mistakes log.', '**5. Recap** (2 minutes): the key points, and new review cards for your deck.', 'The **Outline** button in a lesson lists every slide and task, so you can jump straight to one.'] },
  { id: 'unlock', title: 'How modules unlock', tags: 'locked lock unlock test out padlock module', img: P('map'), body: ['**Modules 1 and 2 are free for everyone.**', 'With a student login, the rest open one at a time: pass a module’s test (70% or more) and the next module opens. Your learning map shows this as diamonds between the stops.', '**Already know a module?** Open it on the map and choose “Take its test now”. Pass, and the module is marked *tested out* and the next one opens.', 'Inside an open module you can take the lessons in any order, and the test whenever you feel ready.', 'Your teacher can also open every module for you from their Admin page.'] },
  { id: 'saved', title: 'Where your progress is saved', tags: 'progress saved backup browser data privacy', body: ['Sheetah has no server. Everything you do is saved **in this browser on this laptop**: lessons, cards, notes, the journal and test results. Nothing is sent anywhere.', 'That means clearing your browser data, or switching laptops, leaves your progress behind. Use **Settings › Your data › Download a backup** now and then, and **Restore a backup** on the other laptop.', 'Signing in for the first time offers to bring along what you did as a visitor.'] },
  { id: 'notquite', title: 'Why a task says “Not quite”', tags: 'wrong check not quite formula hint solution', img: P('lab'), body: ['Sheetah checks what you actually did. For a formula task it checks that the cell **holds a formula**, that it gives the **right answer**, and that it **still works when the numbers change**. That last one catches answers typed in by hand.', 'Read the message under the task: it names the exact problem. Then try **Hint 1, 2 and 3**.', 'After three hints you can open **Show solution**. The task then counts as “done with help” and comes back in your daily review.', 'Open **What Sheetah checked** under a result to see each check.'] },
  { id: 'upload', title: 'Uploading a Real Excel file', tags: 'upload real excel xlsx file error project', img: P('upload'), body: ['1. Press **Practice workbook** to download the starter file.', '2. Open it in Excel (desktop, or the free Excel for the web), do the steps, and save.', '3. Make sure it’s saved as **Excel Workbook (.xlsx)**: File › Save As › Excel Workbook.', '4. Press **Upload my file**. Sheetah checks it at once and lists each check.', '**Common problems:** a .xls or .csv file (save as .xlsx instead), a password-protected workbook (remove the password in File › Info › Protect Workbook), or a file from Numbers or Google Sheets (download it as .xlsx first).'] },
  { id: 'noexcel', title: 'I don’t have Excel', tags: 'no excel licence free web online office', body: ['You don’t need Excel for the lessons: the demos and labs run inside Sheetah.', 'For the optional **Real Excel tasks**, the project and real practice, use the free **Excel for the web** at office.com with a free Microsoft account. It opens .xlsx files in the browser, and almost everything in Excel Foundations works there.', 'The Excel phone apps also open .xlsx files, but a laptop is much easier for learning.'] },
  { id: 'mac', title: 'Mac or Windows keys', tags: 'mac windows keyboard command cmd ctrl option', body: ['Choose your computer in **Settings › Keyboard**. Sheetah then shows the right keys everywhere: Ctrl becomes ⌘ and Alt becomes Option on a Mac.', 'A few Excel shortcuts are different on a Mac, such as **⌘+T** instead of F4 to cycle $ signs. The Shortcuts page lists both.', 'On many laptops the F-keys need the **Fn** key held down.'] },
  { id: 'install', title: 'Install Sheetah on your laptop', tags: 'install app offline desktop icon pwa', body: ['In **Chrome or Edge**, click the install icon at the right end of the address bar (or ⋮ › Install Sheetah). Sheetah then opens in its own window with its own icon, and works offline for lessons you’ve already opened.', 'On a **Mac with Safari**: File › Add to Dock.'] },
  { id: 'move', title: 'Moving to another laptop', tags: 'move another laptop computer backup restore transfer', body: ['1. On the old laptop: **Settings › Your data › Download a backup**. You get a small .json file.', '2. Send the file to yourself (email, WhatsApp, a pen drive).', '3. On the new laptop: open Sheetah, sign in if you have a login, then **Settings › Your data › Restore a backup** and choose the file.'] },
  { id: 'review', title: 'Daily review and cards', tags: 'review cards spaced repetition again hard good easy', img: P('review'), body: ['Each lesson adds review cards to your deck. The daily review brings each card back just before you’d forget it.', 'Read the card, think of the answer, then **Show answer** and rate yourself: **Again** (you didn’t know it), **Hard**, **Good** or **Easy**. Sheetah spaces the next review by your rating.', 'Past mistakes come back in the review too, after 1, 3 and 7 days.', 'Ten to fifteen minutes a day is plenty. Change how many new cards you get per day in Settings › Learning.'] },
  { id: 'mistakes', title: 'The mistakes log', tags: 'mistakes error log wrong cause', body: ['Every wrong lab or quiz answer lands in **Mistakes**, next to the right answer.', 'Tag each one with a cause (didn’t know it yet, careless slip, wrong reference…) to see your patterns.', 'Each mistake returns in your daily review after 1, 3 and 7 days until you’ve got it.'] },
  { id: 'tests', title: 'Tests, the project and the final exam', tags: 'test exam project pass mark retake grade', body: ['**Module tests:** about 25 minutes, 15 questions and 2 practical tasks, pass mark 70%. A failed test lists the lessons to revisit. The retake opens after 24 hours, or straight away once you redo those lessons.', '**Practice tests:** mixed questions from lessons you’ve done, with a target score you set.', '**Course project:** a workbook you build and upload. Pass mark 70%, two submissions.', '**Final exam:** 90 minutes, 40 questions and 6 tasks, pass mark 75%, up to 3 attempts 48 hours apart.', '**Grade:** 60% exam and 40% project. 90% or more earns a distinction.'] },
  { id: 'certificate', title: 'Certificates and verifying them', tags: 'certificate pdf linkedin verify', body: ['Pass the project and the final exam to earn the course certificate. Download it as a PDF or print it, and add it to LinkedIn with one click.', 'Each certificate has an ID and a verify link. Send the ID to your teacher: once they add it to the register, anyone can check it on the verify page.'] },
  { id: 'teacher', title: 'Asking your teacher (and forgotten passwords)', tags: 'teacher help password forgot login contact whatsapp email', body: ['Press **Ask your teacher** (in Help, on the sign-in page, and under tricky tasks). Sheetah writes a message for you with where you are and what went wrong, and opens WhatsApp or email.', '**Forgot your password?** Use the same button on the sign-in page. Your teacher can set a new one for you; there is no self-service reset because Sheetah has no server.'] },
  { id: 'keys', title: 'Keyboard keys in Sheetah', tags: 'keyboard keys shortcuts search fold full screen', body: ['**Ctrl+K** or **/**: search everything.', '**[**: fold or open the side panel.', '**← →**: previous and next slide in a lecture.', '**N**: next part of a lesson.', '**F**: focus mode (full screen) in a lesson. **Esc** leaves it.', '**?**: show this list of keys.'] },
  { id: 'privacy', title: 'What Sheetah stores, and where', tags: 'privacy data stored safe', body: ['Your progress, notes and journal stay in this browser only. Sheetah doesn’t use cookies for tracking, ads or analytics, and no AI is used anywhere in the app.', 'Your teacher sees your progress only if you send it to them yourself (Settings › Your data › Send my progress).'] },
];

// Glossary words in lessons: lightly underlined; hover or tap for the meaning
let GLOSS = [];
export const setGlossary = g => { GLOSS = (g || []).map(([term, def, course, lesson]) => ({ term, key: term.split('(')[0].trim(), def, course, lesson })).filter(x => x.key.length > 2).sort((a, b) => b.key.length - a.key.length); };
export function glossify(root, max = 6) {
  if (!GLOSS.length || !root) return; let n = 0; const used = new Set();
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, { acceptNode: t => t.parentElement.closest('code,kbd,h1,h2,h3,button,a,.gl,.opts,.feedback,label,select,textarea,.xg,figcaption') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT });
  const nodes = []; while (walker.nextNode()) nodes.push(walker.currentNode);
  for (const node of nodes) {
    if (n >= max) break;
    for (const g of GLOSS) {
      if (used.has(g.key)) continue; const re = new RegExp(`\\b(${g.key.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})s?\\b`, 'i'); const m = re.exec(node.data); if (!m) continue;
      const span = document.createElement('span'); span.className = 'gl'; span.tabIndex = 0; span.dataset.g = g.key; span.textContent = m[0];
      const after = node.splitText(m.index); after.data = after.data.slice(m[0].length); node.parentNode.insertBefore(span, after); used.add(g.key); n++; break;
    }
  }
}
export function showGloss(el) {
  const g = GLOSS.find(x => x.key === el.dataset.g); if (!g) return;
  popover(el, `<b>${esc(g.term)}</b><p>${md(g.def)}</p>${g.lesson ? `<a href="#/lesson/${g.course}/${g.lesson}">Taught in Module ${g.lesson.split('-')[0]} · Lesson ${g.lesson.split('-')[1]} →</a>` : ''}`);
}

// "Ask your teacher": a ready message for WhatsApp or email
export function askTeacher({ where = '', what = '', subject = 'Help with Sheetah' } = {}) {
  const T = teacher(); const name = (session.s && session.s.name) || store.d.name || ''; const email = session.s ? session.s.email : '';
  const text = `Hi${T.name ? ' ' + T.name : ''}, this is ${name || 'a Sheetah learner'}${email ? ` (${email})` : ''}.\n${where ? `I'm on: ${where}\n` : ''}${what || 'I’m stuck and could use some help.'}`;
  const m = modal(`<h2>Ask your teacher</h2><p>Sheetah wrote a message for you. Change it if you like, then send it.</p><textarea id="atxt" rows="7">${esc(text)}</textarea>
   <div class="row end">${T.whatsapp ? '<a class="btn" id="awa" target="_blank" rel="noopener">Send on WhatsApp</a>' : ''}${T.email ? '<a class="btn ghost" id="aem">Send by email</a>' : ''}<button class="btn ghost" id="acp">Copy message</button></div>
   ${!T.whatsapp && !T.email ? '<p class="muted small">Your teacher hasn’t added contact details yet. Copy the message and send it the way you usually reach them.</p>' : ''}`);
  const cur = () => $('#atxt', m.el).value;
  const wa = $('#awa', m.el); if (wa) wa.onclick = () => { wa.href = waLink(cur(), T.whatsapp); };
  const em = $('#aem', m.el); if (em) em.onclick = () => { em.href = mailLink(T.email, subject, cur()); };
  $('#acp', m.el).onclick = () => { navigator.clipboard && navigator.clipboard.writeText(cur()); toast('Message copied'); };
}

// pages and settings that search can find (lessons, topics, functions, shortcuts and terms are added from content)
export const PAGES = [
  ['Home', '#/home', 'start today plan continue'], ['My courses', '#/courses', 'courses catalogue foundations formulas'], ['Learning map: Excel Foundations', '#/course/c1', 'map modules path where am i'],
  ['Daily review', '#/review', 'cards spaced repetition'], ['Tests', '#/tests', 'practice mock test score'], ['Placement check', '#/placement/c1', 'placement skip level test out'],
  ['Playground', '#/playground', 'sheet try formula open file'], ['Weekly challenge', '#/challenge', 'challenge case practice'], ['Interview prep', '#/interview', 'interview questions job'],
  ['Function library', '#/library', 'functions sum average'], ['Shortcuts', '#/shortcuts', 'keys keyboard trainer'], ['Topic index', '#/topics', 'topics a-z index'],
  ['Study plan', '#/planner', 'plan schedule days calendar reminder'], ['Journal', '#/journal', 'diary check-in mood weekly'], ['Mistakes', '#/errors', 'error log wrong'], ['Insights', '#/insights', 'charts stats mastery calendar'],
  ['Certificates', '#/certificates', 'certificate pdf linkedin'], ['My notes', '#/notes', 'notes word'], ['Saved', '#/saved', 'bookmarks saved'], ['Downloads', '#/downloads', 'files workbooks notes word excel'],
  ['Help centre', '#/help', 'help faq questions'], ['How Sheetah works', '#/how', 'guide video tour how'], ['Student sign-in', '#/login', 'login password'],
  ['Settings: appearance (theme, colours, text size)', '#/settings/look', 'dark light theme colour text size font contrast'], ['Settings: layout (side panel, spacing, mascot)', '#/settings/layout', 'side panel compact mascot'],
  ['Settings: learning (goal, days, reminders, tone)', '#/settings/learning', 'goal reminders tone days review cards hints'], ['Settings: keyboard (Windows or Mac)', '#/settings/keyboard', 'mac windows keys'],
  ['Settings: profile and language', '#/settings/profile', 'name avatar language hindi'], ['Settings: your data (backup, move, reset)', '#/settings/data', 'backup restore move laptop reset send progress'],
  ['Run setup again', '#/setup', 'setup wizard start again'], ['Take the tour', '#/tour', 'tour guide'],
];
