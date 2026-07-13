Multiple RSVP match support

This version supports more than one RSVP group containing the same guest first and last name.

Behavior:
- If no match is found, the guest may start a new RSVP.
- If one group matches, the normal confirmation appears.
- If multiple groups match, the page lists every matching group and the names in each party.
- The guest selects "This is my RSVP" to open the correct household.
- Updates continue to use the permanent Household ID, so only the chosen group is modified.

Important:
Replace your Apps Script Code.gs with the included version, save, and redeploy the Web App.
If redeployment creates a new /exec URL, update RSVP_ENDPOINT in rsvp.html.
