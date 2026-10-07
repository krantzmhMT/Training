# Training Data sync

`Code.gs` connects three things through one private Google Sheet in your Drive:

- **Morning Coach → Sheet.** Each morning check is saved to the *Mornings* tab.
- **Garmin → intervals.icu → Sheet.** Every 3 hours, wellness data goes to the *Wellness* tab and runs go to the *Activities* tab.
- **Coach → app.** The newest row in the *Coach notes* tab appears at the top of Morning Coach.

The script itself lives here in the repo. The secrets do not: the intervals.icu API key and the app token are kept in the script's Script Properties on Google's side.

## Setup

1. Open the **Training Data (Morning Coach)** sheet. Go to **Extensions → Apps Script**.
2. Delete the placeholder code, paste in all of `Code.gs`, and click **Save**.
3. Click **Deploy → New deployment**. Click the gear icon and choose **Web app**. Set:
   - **Execute as:** Me
   - **Who has access:** Anyone

   Click **Deploy**, then **Authorize access**. Google will warn that the app isn't verified. That's expected for your own script: click **Advanced → Go to … (unsafe)**, then **Allow**. Copy the **Web app URL**; it ends in `/exec`.
4. Reload the Sheet. A **Training sync** menu appears.
   - **1 · Connect intervals.icu:** paste your API key (intervals.icu → Settings → Developer Settings). This pulls the first 90 days.
   - **2 · Show app connection link:** paste the Web app URL from **Deploy → Manage deployments** when it asks, then copy the link it shows.
5. In Morning Coach, tap ⚙ and paste the link into **Sync link**.

## If you change Code.gs later

Paste the new version into the editor and save. Then go to **Deploy → Manage deployments**, edit the existing deployment (pencil icon), choose **Version: New version**, and click **Deploy**. This keeps the same URL, so the app needs no change. Choosing **New deployment** instead creates a new URL.
