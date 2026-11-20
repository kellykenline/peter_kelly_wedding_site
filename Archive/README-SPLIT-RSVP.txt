Peter & Kelly Website — Split RSVP Version

Files:
- index.html: main wedding website. The RSVP section now links to rsvp.html.
- rsvp.html: dedicated RSVP management page. No countdown appears on the RSVP page.
- assets/: image files required by the site.
- Code.gs: Google Apps Script backend for RSVP lookup/editing.

Important:
Replace Apps Script Code.gs with the included Code.gs, save, and redeploy.

RSVP behavior:
- Guests search by any submitted guest's first name + last name.
- If a match is found, the full group RSVP opens for editing.
- If no match is found, guests can start a new RSVP.
- Each guest now collects and sends:
  - Guest First Name
  - Guest Last Name
  - Individual Attendance
  - Adult/Child
  - Meal Selection, when attending
  - Dietary Restrictions / Other Comments

The spreadsheet receives Guest First Name and Guest Last Name as separate columns for every new RSVP and every update.
