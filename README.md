# Outlook Calendar Privacy Tools (VBA)

This repository contains two simple Microsoft Outlook VBA macros to help you manage the visibility of your calendar events by toggling the **"Private"** flag in bulk.
- Mask confidential meetings from shared calendars.
- Clean up your calendar visibility without disabling access entirely.

These scripts are for anyone who shares their Outlook calendar with colleagues or assistants and wants to block the viewing of details for events within certain specified time period, without affecting meeting functionality or availability.

---

## 📜 Scripts

### 🔒 `Set-All-To-Private.bas`

**Purpose**:  
Marks all calendar events within a specified date range as **Private**, preventing others with "View All Details" access from seeing the event subject, location, and body content.

**How it works**:
- Scans your default Outlook calendar for events from 180 days ago to 180 days in the future.
- Sets the `Sensitivity` property of each `AppointmentItem` to `olPrivate`.
- Skips inaccessible items automatically.

> 💡 Marking a meeting as "Private" does **not** affect your ability to accept/decline invites or receive updates.

---

### 🔓 `Unset-All-To-Private.bas`

**Purpose**:  
Reverses the above change. Resets all `AppointmentItem` objects within a 400-day window to `Sensitivity = olNormal`, making them visible again to others with access to your calendar.

---

## 🛠️ Usage Instructions

1. Open Outlook and press `ALT + F11` to open the **VBA editor**.
2. Go to `Insert > Module`, and paste in one of the scripts.
3. Close the editor and run the macro via `ALT + F8`, selecting the relevant script.
4. Outlook will display a confirmation message when it’s done.

> 🧪 Recommended: Test on a few dummy events or a narrowed date range before running broadly.

---

## ⚠️ Notes

- These macros only work in **Windows desktop Outlook** (not Outlook Web App).
- They operate on your **default calendar only**.
- The macros do not affect recurring appointments differently — each occurrence is handled individually.
- Events outside the configured date range are **left untouched**.


## 🔐 Privacy Impact

- Marking an event as "Private":
  - Hides details from users with "View All Details" permission.
  - Still allows updates from meeting organizers.
  - Still allows you to RSVP to meetings.

---

## ✅ License

MIT License — Feel free to modify and share.

---
