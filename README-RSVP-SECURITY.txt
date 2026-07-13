RSVP security and reliability update

Implemented:
- Server-side validation for all submitted fields
- Maximum lengths for names and comments
- Allowlisted attendance, guest type, meal, and event values
- Spreadsheet formula-injection protection
- Safe HTML escaping when saved comments are displayed
- Script locking during delete-and-rewrite updates
- Hidden honeypot bot field
- Minimum form completion time
- Basic per-name rate limiting
- Confirmed-save verification before displaying the success message
- Existing duplicate-name household selection remains supported

Deployment steps:
1. Replace the Apps Script Code.gs with the included Code.gs.
2. Save the Apps Script project.
3. Redeploy the Web App.
4. Send the new /exec URL so it can be placed in rsvp.html.
5. Test a new RSVP and an RSVP update before publishing.
