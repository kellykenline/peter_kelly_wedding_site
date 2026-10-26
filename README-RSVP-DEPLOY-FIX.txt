RSVP Apps Script fix

1. In your RSVP Google Sheet, open Extensions > Apps Script.
2. Replace the entire contents of Code.gs with the Code.gs from this ZIP.
3. Save.
4. Choose Deploy > Manage deployments, edit the existing Web App deployment, select New version, then Deploy.
5. Verify the existing deployment URL remains the same as the API_URL in rsvp.html.
6. Test the RSVP page with a real household code.

The updated Code.gs adds doGet for browser lookups while preserving doPost for the website.
