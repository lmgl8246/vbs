# Holidays_2026.vbs

`Holidays_2026.vbs` is a Windows VBScript that adds 2026 City of Austin holidays to your default Microsoft Outlook Calendar.

## What it does

When you run the script:

1. You are prompted to confirm whether to continue.
2. The script opens Outlook and connects to your default calendar.
3. It creates one all-day appointment for each configured holiday date.
4. Each holiday is saved with:
   - **Reminder enabled**
   - **Reminder set 7 days before** (10,080 minutes)
   - **Busy status set to Out of Office**

After all entries are created, you get a completion message.

## Included 2026 holidays

- January 1, 2026 — New Year's Day
- January 19, 2026 — Martin Luther King Day
- February 16, 2026 — President's Day
- May 25, 2026 — Memorial Day
- June 19, 2026 — Juneteenth
- July 3, 2026 — Independence Day Observed
- September 7, 2026 — Labor Day
- November 11, 2026 — Veterans Day
- November 26, 2026 — Thanksgiving Day
- November 27, 2026 — Thanksgiving Friday
- December 24, 2026 — Christmas Eve
- December 25, 2026 — Christmas Day

## Requirements

- Windows with **Windows Script Host** enabled
- Desktop **Microsoft Outlook** installed and configured
- Access to your default Outlook profile/calendar

## How to run

1. Save/keep `Holidays_2026.vbs` on your machine.
2. Close Outlook if you want to ensure a clean start (optional).
3. Double-click `Holidays_2026.vbs`, or run from Command Prompt:

```bat
cscript //nologo Holidays_2026.vbs
```

4. Click **Yes** on the confirmation dialog.

## Important notes

- The script does **not** check for duplicates. Running it multiple times will create duplicate holiday entries.
- Holidays are created in the **default Outlook calendar** for the active profile.
- Entries are set as all-day events from `12:00 AM` to `11:59 PM` on each date.

## Safety recommendation

Before first run, consider exporting your calendar or creating a temporary backup in Outlook.

## Customization

To change or add holidays, edit the dictionary entries in `Holidays_2026.vbs`:

```vbscript
objDictionary.Add "January 1, 2026", "New Year's Day"
```

Use the format:

- **Date key:** `Month Day, Year`
- **Value:** Holiday/event name
