# Email Automation

A desktop app for sending personalized bulk emails. Write one email template, import a list of recipients, and the app sends each person a customized version with their own name, company, promo code, or any other data you provide.

---

## What It Does

- **Personalized emails at scale** — Write your email once using `{{variable}}` placeholders. Each recipient gets their own version with the placeholders swapped out for their real data.
- **Rich text formatting** — Bold, italic, underline, and hyperlinks, just like Gmail.
- **CSV import** — Import a spreadsheet of recipients. Column headers become your variable names.
- **Built-in recipient builder** — No spreadsheet? Build your recipient list directly inside the app, then export it or send straight away.
- **Preview before sending** — See exactly what each recipient's email will look like before anything is sent.
- **Works with Gmail, Outlook, and Yahoo** — Pick your provider from a dropdown. No technical SMTP knowledge needed.

---

## How Variables Work

In your email you write `{{variable_name}}`. In your CSV (or builder), each column header is a variable name. When the email is sent, the placeholder is replaced with that person's data.

**Example email:**
```
Subject: Hey {{first_name}}, an offer just for you!

Hi {{first_name}},

We have a special deal for everyone at {{company}}.
Use code {{promo_code}} at checkout for 20% off.
```

**Example CSV:**
```
email,first_name,company,promo_code
alice@example.com,Alice,Acme Corp,SAVE20
bob@example.com,Bob,Globex,SAVE15
```

Alice gets: *"Hey Alice, an offer just for you!"* with her company and promo code.
Bob gets his own personalized version.

---

## Requirements

- **Python 3.10 or newer** — download from [python.org](https://python.org)
- No third-party packages required — uses Python's standard library only

---

## How to Run

**1. Install Python** (if you haven't already)

Go to [python.org](https://python.org), download the installer, and make sure to check **"Add Python to PATH"** during installation.

**2. Download or clone this project**

If you downloaded a ZIP, extract it somewhere on your computer.

**3. Open a terminal in the project folder**

- On Windows: open the folder, click the address bar, type `cmd`, press Enter
- Or open Command Prompt / PowerShell and navigate to the folder:
  ```
  cd path\to\emailAutomation
  ```

**4. Run the app**

```
python main.py
```

The app window will open.

---

## Setup (First Time)

### SMTP Setup tab
This tells the app which email account to send from.

- Choose your **Email Provider** (Gmail, Outlook, or Yahoo)
- Enter your **Username** (your full email address)
- Enter your **Password** — see the note for your provider:
  - **Gmail**: You must use an **App Password**, not your regular password. Go to [myaccount.google.com](https://myaccount.google.com) → Security → App passwords to create one. (Requires 2-Step Verification to be enabled.)
  - **Outlook**: Your regular password works.
  - **Yahoo**: Requires an App Password. Go to Yahoo Account Security → Generate app password.
- Click **Save Settings**, then **Test Connection** to make sure everything works.

---

## Sending Emails

### Step 1 — Compose tab
Write your subject and body. Use `{{variable_name}}` anywhere you want personalized content. The app will detect your variables and show them as you type.

Use the toolbar to **bold**, *italic*, underline, or insert links. Right-click in the body for cut/copy/paste.

### Step 2 — Add recipients (pick one)

**Option A — Import CSV tab**

Click **Import CSV** and select your `.csv` file. The first row must be column headers. One column must be named `email` (or select it from the dropdown). Click **Show CSV Format Example** if you need help with the format.

**Option B — Build Recipients tab**

Build your list directly in the app:
- Click **+ Add Column** to define your columns (start with `email`)
- Click **+ Add Row** to add people
- **Double-click any cell** to edit it
- Click **Use for Sending** when done

### Step 3 — Preview & Send tab
- Use the **Preview for** dropdown to see exactly what each person's email will look like
- Click **Send All Emails** to start sending
- Click **Stop** at any time to cancel mid-batch
- The send log shows the result for every email

---

## Tips

- The app sends one email at a time with a short delay between each (300 ms by default) to avoid being flagged as spam. You can adjust this on the SMTP Setup tab.
- `smtp_config.json` is created automatically to save your settings. **Do not share this file** — it contains your password.
- If Gmail blocks the connection, make sure you are using an App Password and that 2-Step Verification is enabled on your Google account.
