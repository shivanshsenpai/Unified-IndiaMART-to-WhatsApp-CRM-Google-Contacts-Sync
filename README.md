# Unified-IndiaMART-to-WhatsApp-CRM-Google-Contacts-Sync
A robust, thread-safe Google Apps Script architecture that automates the entire lead pipeline. It directly ingests leads from the IndiaMART API, cleans and deduplicates them in Google Sheets, categorizes them using regex rules, dynamically formats names to bypass WhatsApp's 256-member broadcast limit, and automatically syncs them to Google Contacts
# Unified IndiaMART to WhatsApp CRM & Google Contacts Sync

This project transforms a standard Google Sheet into a fully automated CRM pipeline. It bridges the gap between raw lead generation (IndiaMART) and actionable outreach (WhatsApp and Google Contacts), completely bypassing Google Apps Script’s native execution timeouts and API rate limits through advanced state management and thread-safe locking.

## ✨ Core Features

* **Automated Lead Ingestion:** Pulls directly from the IndiaMART API. Auto-splits large date ranges into 7-day chunks to respect API limits and prevents infinite loops using payload fingerprinting.
* **Smart Deduplication & Sanitization:** Normalizes Indian phone numbers and checks them against an in-memory hash map to instantly drop duplicates before they hit your database.
* **Regex Categorization:** Reads the product query and assigns leads to specific categories based on custom Regex rules defined in your sheet.
* **WhatsApp Broadcast Optimization:** Automatically appends grouping tags (e.g., `SHO_G1_C145_Name`). Once a group hits 200 members, it automatically rolls over to `G2` to stay safely under WhatsApp's 256-member broadcast limit.
* **Thread-Safe Contact Sync:** Uses Google's `LockService` and a pre-claim status ("Syncing...") to safely push leads to Google Contacts in the background without creating duplicates if multiple triggers fire simultaneously.
* **One-Click CSV Exports:** Generates clean CSV files grouped by category and saves them directly to your Google Drive for easy bulk messaging.

## ⚙️ Setup Instructions

### 1. Sheet Preparation
Create a new Google Sheet and ensure it has the following four tabs (exact naming matters):
1. `customer`: The raw landing zone for incoming API leads.
2. `Processed_data`: The cleaned, deduplicated, and categorized master list.
3. `Config`: Your rule engine for categorizing leads.
4. `Logs`: The system's automatic diagnostic log.

### 2. The Config Sheet Structure
In your `Config` tab, set up the following columns:
* **A (Category):** e.g., `Shoes`
* **B (Initial):** e.g., `SHO`
* **C (Keywords):** Pipe-separated regex keywords, e.g., `sneaker|boots|heels`
* **D (Current Group):** Start at `1` (The script updates this automatically)
* **E (Current Count):** Start at `0` (The script updates this automatically)

### 3. Script Installation
1. Open your Google Sheet > **Extensions** > **Apps Script**.
2. Delete any code in the editor and paste the generalized code provided below.
3. Replace `YOUR_INDIAMART_API_KEY_HERE` with your actual IndiaMART CRM key.
4. Save and refresh your spreadsheet. A new custom menu called **IndiaMART + WhatsApp CRM** will appear.

### 4. Permissions
The first time you run any function from the custom menu, Google will ask for permissions to access your Sheets, Drive, and Google Contacts. Allow these permissions.

## 🚀 Usage
* **Initialize Automation:** Run `Initialize CRM Automation` from the custom menu first. This sets up your on-edit and 5-minute background sync triggers.
* **Pull Leads:** Open the `Lead Puller`, select your dates, and hit start. The script handles the rest.
* **Set & Forget:** Use `Setup Recurring Import` to have the script automatically look backward *N* days and pull leads continuously.
