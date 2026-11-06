# PQL office hours

Booking page for PQL office hours. PQLs pick an open 2-hour slot and add their name. The Live now tab shows who is on the Google Meet call right now.

**Live site:** https://harman-arneja.github.io/pql-office-hours/

## How it works

- `index.html` is the whole page. It is hosted on GitHub Pages.
- The data lives in a private Google Sheet (tabs: Settings, PQLs, Bookings, Live).
- `Code.gs` is the Apps Script bound to that sheet, deployed as a web app with access set to Anyone. The page calls it with plain GET requests (`?fn=getState`, `?fn=book`, `?fn=unbook`, `?fn=adminLogin`, `?fn=adminSave`).
- The admin password is stored in the sheet and checked by the script. It is never in this repo.

## Changing things

- Slot times, slot length, seats per slot, slots per PQL, the meeting link and the PQL list: use the Admin tab on the site, or edit the Settings and PQLs tabs in the sheet.
- After changing `Code.gs`, paste it into the Apps Script project and deploy a new version (Deploy → Manage deployments → Edit → New version). The web app URL stays the same.
