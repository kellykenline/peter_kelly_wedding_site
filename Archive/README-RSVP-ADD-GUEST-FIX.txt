RSVP update fix

This version fixes editing an existing RSVP and increasing the guest count.

Changes:
- Existing guest information is preserved when the number of guests changes.
- New guest cards are added without clearing the people already entered.
- The submitted guest count is calculated from the actual guest records.
- Code.gs validates every guest and writes all rows to the spreadsheet in one batch.

Important:
Replace the Apps Script Code.gs with the included version, save it, and redeploy the Web App.
If the deployment creates a new /exec URL, update RSVP_ENDPOINT in rsvp.html.
