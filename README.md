# Philadelphia PA Address Search Tool

This version does **not** use Chrome or ChromeDriver.

It avoids GitHub Actions browser errors by using:

1. Philadelphia AIS API to convert address input to OPA account number.
2. `https://property.phila.gov/?p=<OPA_NUMBER>` to scrape the property page.
3. Excel output using OpenPyXL.

## Files

```text
Philadelphia_PA_AddressSearch_Tool.py
PA.txt
requirements.txt
README.md
.gitignore
.github/workflows/run_philadelphia_scraper.yml
```

## Run locally

```powershell
pip install -r requirements.txt
python Philadelphia_PA_AddressSearch_Tool.py
```

## Run in GitHub

```text
GitHub repo → Actions → Run Philadelphia PA Address Search Tool → Run workflow
```

## Schedule

The workflow runs every 6 hours:

```yaml
cron: "0 */6 * * *"
```

India time:

```text
5:30 AM
11:30 AM
5:30 PM
11:30 PM
```

## Optional AIS key

If the City later requires an AIS key, add this GitHub repository secret:

```text
AIS_GATEKEEPER_KEY
```

The code will use it automatically if present.
