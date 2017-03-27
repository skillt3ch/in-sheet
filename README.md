# In-Sheet ReFormatter
Scrapes inbox for web service requests and pulls all the data.
Using a HTML template (*template.html*), a new HTML document is populated with the scraped data from the email.
The HTML is then converted to PDF using **weasyprint**.

## Instructions:

### Requirements

Run the following:
```bash
pip install -r requirements.txt
```

### Usage

```bash
insheet.py [-u <username>] [-j <job number>] [-f <folder>]
```

You can run this script with optional arguments.
The username and job number are both set to initialised values.
Username is just the first half of your email address **username**@company.com

If no job number is given, all emails will be scraped and returned.
If job number is provided, only the email containing that job number will be returned.

* Options for -f:
	* s = "Service Requests" (default)
	* sd = "Service Requests DONE"
	* i = "INBOX"

e.g.

```bash
insheet.py -u johnsmith -j 703803 -f sd
```

## To Do:

* [X] Add command line arguments to enable extra functionality:
	* [X] Create new in-sheet for a single job number
	* [X] Choose which mailbox to search (*'Service Requests' / 'Service Requests DONE', 'INBOX', etc.*)
	* [X] Choose which email account to log in as (currently hardcoded)
* [X] Add comments to code
* [X] Add error handling (e.g. incorrect password, errors with scraping, errors if job number not found, etc.)
