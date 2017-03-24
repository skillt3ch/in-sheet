# In-Sheet ReFormatter
Scrapes inbox for web service requests and pulls all the data.
Using a HTML template (*template.html*), a new HTML document is populated with the scraped data from the email.
The HTML is then converted to PDF using **weasyprint**.

## Instructions:

### Usage

```bash
insheet.py [-u <username>] [-j <job number>]
```

You can run this script with optional arguments.
The username and job number are both set to initialised values.
Username is just the first half of your email address **username**@company.com

If no job number is given, all emails will be scraped and returned.
If job number is provided, only the email containing that job number will be returned.

e.g.

```bash
insheet.py -u johnsmith -j 703803
```

## To Do:

* Add command line arguments to enable extra functionality:
	* ~~Create new in-sheet for a single job number~~
	* Choose which mailbox to search (*'Service Requests' / 'Service Requests DONE', 'INBOX', etc.*)
	* ~~Choose which email account to log in as (currently hardcoded)~~
* ~~Add comments to code~~
* Add error handling (e.g. incorrect password, errors with scraping, errors if job number not found, etc.)
