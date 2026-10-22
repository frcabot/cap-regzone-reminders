# CAP Reg Zone reminder emails

Automatic, personal reminder emails for a Civil Air Patrol activity, built from **Registration Zone (Reg Zone)** exports. Use it for an encampment, an academy, or any activity that registers through Reg Zone.

You save the Reg Zone exports to a Google Drive folder. The script then emails each person **only what they still owe**, with a due date for each item. Anyone with nothing left to do gets no email.

| Email | Who gets it | How often |
|---|---|---|
| **To-do list** | Staff, and students who are **slotted** | About twice a week, until everything is done |
| **Waitlist update** | Alternates (registered, not slotted) | About once a week |
| **Unit commander list** | Each unit commander whose members are registered | About once a week |

- **Parents** are included only for cadets **under 18**.
- **The unit commander** is copied on a member's email only while that member is missing **unit approval**.
- **Every email** says up top that no reply is needed unless there is a question.

Everything the script plans to send is listed on a sheet tab first, so you can check it before anything goes out.

---

## What goes on the to-do list

Each item can be turned off, and each has its own due date. You set these in `ITEMS` at the top of the script.

| Item | Who | Where the script reads it |
|---|---|---|
| Unit commander approval | Everyone | Registration export (`UnitApproved`) |
| Parent/guardian approval, with how-to steps (linked or not linked) | Cadets under 18 | `ParentApproved`, `LinkedToParentAccount` |
| Link a parent account in eServices | Cadets under 18 whose parent approved but has no linked account | `LinkedToParentAccount` |
| Add a parent email in eServices | Cadets under 18 with no primary parent email | `CadetParentEmailPrimary` |
| Promotion to C/Amn | Slotted C/AB students (alternates: as soon as possible) | `Rank` = `CADET` |
| Payment | Everyone, with separate student, cadet staff and senior fees | `AmountPaid`, `PaidInFull` |
| RST, completed on or after a date you choose | Staff | Prerequisite report (`Training`) |
| CPPT: complete, or renew if it runs out before the last day | Everyone 18 or older by the last day | `CPPTExpiration` (both reports) |
| Medical profile current in CAP Health | Everyone | Prerequisite report |
| Level I | Senior staff | Prerequisite report (`Member Level`) |
| Your own items (forms, photos, anything) | The CAPIDs you list | **Extra items** tab |

**First-time CPPT for cadets.** The email explains that it needs the online course **and** a summary conversation with the squadron commander. A renewal needs only the course.

**Waitlist updates.** Alternates see:
- their place in line, either in one line or in separate male and female lines (for all-male and all-female flights);
- their chances, based on how many slotted people usually drop out;
- what keeps them from being slotted, if anything;
- how to leave the waitlist.

Place numbers count only alternates who can be slotted right now.

**Unit commander list.** A formatted table with one row per member and one column per item. A blank cell means nothing is needed. The first row shows how many still need each item. Below it, every alternate in the unit is listed with their place in line.

---

## What you need

- A Google account that can send email. A Google Workspace account sends more emails per day than a personal Gmail account.
- The Reg Zone exports for your activity, saved as `.xlsx` files in one Google Drive folder:
  - **One registration export per Reg Zone activity** (for example staff, students, and an advanced training squadron).
  - *(Optional, recommended)* **The Attendance Prerequisite report** for each activity. It adds RST dates, medical profile status, Level I, and the real CPPT expiration date.
- Save each fresh export under the **same file name** every time, replacing the old file. The script finds files by name.

### Columns the script reads

From the registration export: `RegistrantsCAPID`, `Rank`, `NameFirst`, `NameLast`, `Gender`, `Age`, `AgeAtEventStart`, `AgeAtEventEnd`, `MbrType`, `Email`, `RegistrationStatus`, `RegistrationID`, `SlotScore`, `UnitApproved`, `ParentApproved`, `LinkedToParentAccount`, `CadetParentEmailPrimary`, `CadetParentEmailSecondary`, `CadetParentEmailEmergency`, `UnitCCName`, `UnitCCEmail`, `Unit`, `Wing`, `AmountPaid`, `PaidInFull`, `CPPTExpiration`, `Addr1`, `EmergencyContactNumber`.

From the prerequisite report: `RegistrantsCAPID`, `PrereqType`, `MeetsRequirement`, `CompletionDate`, `CPPTExpiration`.

`Addr1` and `EmergencyContactNumber` are used only to find a brother's or sister's parent email when a cadet has none. They are never written to the sheet.

---

## Setup

1. **Make a Google Sheet** for the activity.
2. Open **Extensions > Apps Script**. Delete the sample code, paste in all of `RegZoneReminders.gs`, and save.
3. **Fill in the settings** at the top of the script:
   - `CONFIG`:
     - the activity name, location, dates and last day (`YYYY-MM-DD`);
     - your name, title and email;
     - who to contact with questions;
     - the Drive folder ID;
     - one entry in `GROUPS` per Reg Zone export (see below).
   - `ITEMS`: due dates (`'YYYY-MM-DD'`), `''` for "as soon as possible", or `false` to leave an item out.
   - `FEES`, `RST_FROM`, and optionally `PAYMENT_NOTE`, for example to tell applicants for financial assistance not to pay yet.
   - Optional: `WAITLIST` texts, `STUDENT_NOTE`, `GUIDES` (picture guides), and `CLOSINGS` (the thank-you lines).

   The folder ID is the long code at the end of the folder's web address, after `/folders/`.
4. **Run `setup`.** Pick `setup` in the function list at the top of the editor, click **Run**, and approve the permissions Google asks for.
   - Setup refuses to start while any setting still has its example value.
   - It creates the **Extra items** tab and reports whether every export was found.
5. **Reload the spreadsheet.** A **Reg Zone reminders** menu appears.
6. **Preview** with **Reg Zone reminders > Preview**. The **Reminder emails** tab lists every person, what they still owe (or their place in line), who the email goes to, and whether it is due.
7. **Send test emails.** Five sample emails go to you, and to anyone in `TEST_CC`, using real data. Nobody else is emailed.
8. **Send** with **Send to EVERYONE now**, or one kind or one group at a time. Then, if you like, **Turn ON automatic emails**.

### Groups (`CONFIG.GROUPS`)

```js
GROUPS: [
  { name: 'Staff', type: 'staff', file: 'STAFF_latest.xlsx', prereqFile: 'STAFF_prereq_latest.xlsx' },
  { name: 'Encampment', type: 'students', file: 'STUDENTS_latest.xlsx', prereqFile: 'STUDENTS_prereq_latest.xlsx',
    lines: 'sex', order: 'score' },
  { name: 'ATS', type: 'students', file: 'ATS_latest.xlsx', prereqFile: '', roomForEveryone: true }
]
```

| Setting | Meaning |
|---|---|
| `type: 'staff'` | Everyone registered gets a to-do list. |
| `type: 'students'` | Only **Slotted** registrations get a to-do list. Everyone else is an alternate. |
| `lines: 'sex'` / `'one'` | Separate male and female waitlist lines, or one line. |
| `order: 'score'` / `'signup'` | Line order: Reg Zone `SlotScore` (highest first, then sign-up order), or sign-up order only. |
| `roomForEveryone: true` | No place numbers. An alternate is told they will be slotted as soon as their approvals are done. |
| `prereqFile` | The prerequisite report for this group. Leave it blank if you don't use it; RST and the medical profile are then skipped. |

### Extra items tab

Use this for anything Reg Zone doesn't track, such as a signed form or a photo. Add one row per person and item:

| CAPID | Item (short title) | Details (what to do) | Due (YYYY-MM-DD) | Done? |
|---|---|---|---|---|
| 123456 | Photo release | Return the signed photo release form to the admin officer. | 2026-11-05 | |

The item shows on that person's to-do list (and in the commander table, under "Other") until you type **Yes** in **Done?**.

---

## Day to day

1. Save fresh Reg Zone exports to the folder, under the same file names.
2. Automatic runs (if on) send whatever is due, and email you a summary.
   - A run **stops and warns you** instead of sending if an export is missing, or if the newest exports are older than `MAX_DATA_AGE_HOURS`. People are never reminded about something they finished days ago.

| Menu item | What it does |
|---|---|
| Preview | Fills the Reminder emails tab and counts what would be sent. Sends nothing. |
| Send test emails | Five samples, to you and `TEST_CC` only |
| Send to EVERYONE now / to-do lists / waitlist updates / unit commander lists | Sends by hand. Anyone emailed in the last 12 hours is skipped. |
| Send to ONE group now | To-do lists and waitlist updates for one group only |
| Turn ON / OFF automatic emails | Runs on `RUN_DAYS` at about `HOUR`. To-do lists never twice within `TODO_EVERY_DAYS`; waitlist and commander lists about once a week. |
| Automation status | Shows whether automatic emails are on and which Google account sends them |
| Run automation from THIS account | Sign in as the account that should send the emails (for example a Workspace account), then run this |

**Daily email limits.** If Gmail's daily limit runs low, the remaining emails wait on an **Email queue** tab. They go out automatically within a few hours.

**Google's 6-minute limit.** If a send takes longer, the script stops after 4 minutes and continues by itself 1 minute later.

**Per-person adjustments** are in `ADJUST`:

| Setting | Use it to |
|---|---|
| `NO_EMAIL` | Never email someone |
| `SKIP_ITEMS` | Leave out an item for one person |
| `SENIOR_PRICE` | Charge a cadet in a senior role the senior fee |
| `PAID_OVERRIDE` | Fix a payment amount Reg Zone gets wrong |
| `LEVEL1_PENDING` | Remind a senior who registered before finishing Level I |

---

## Privacy

- **What the sheet shows:** names, CAPIDs, email addresses, what each person still owes, and when they were emailed. It never shows dates of birth, phone numbers, addresses or medical information. Medical profile status is only "current" or "not current".
- **Where the exports stay:** in your Drive folder. Share that folder, and the spreadsheet, only with people who run the activity.
- **What reviewers see:** test emails contain one real person's to-do list or waitlist status, so keep `TEST_CC` to staff who would see that anyway.

---

## Notes and limits

- Reg Zone reports can lag by a day or two. The emails say so, and tell people that finished items drop off on their own.
- **Payment.** For seniors, the amount paid is compared with the fee, because Reg Zone's `PaidInFull` can show Yes at $0 for part-time staff. For cadets, `PaidInFull` also counts.
- **CPPT.** Reg Zone's registration export says "Needs CPPT" even when CPPT is current but expires before the last day. The prerequisite report keeps the real date, so the email can say "renew" instead of "complete".
- **Parent emails.** If a cadet under 18 has no parent email, the script uses the emergency parent email, or a brother's or sister's parent email (same last name and same address or emergency phone). Cadets still without one are listed in the run summary.
- The script is plain Google Apps Script with no outside services. It reads `.xlsx` files itself, so the exports don't need to be converted.
- **Base access.** For base access forms and graduation guest lists at a military installation, see the companion tool [cap-base-access-form](https://github.com/frcabot/cap-base-access-form).
