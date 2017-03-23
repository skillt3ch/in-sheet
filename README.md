# In-Sheet ReFormatter
Scrapes inbox for web service requests and pulls all the data.
Using a HTML template (*template.html*), a new HTML document is populated with the scraped data from the email.
The HTML is then converted to PDF using **weasyprint**.

## Todo:

* Add command line arguments to enable extra functionality:
	* Create new in-sheet for a single job number
	* Choose which mailbox to search (*'Service Requests' / 'Service Requests DONE', 'INBOX', etc.*)
	* Choose which email account to log in as (currently hardcoded)
* Add comments to code
* Add error handling (e.g. incorrect password, errors with scraping, errors if job number not found, etc.)
