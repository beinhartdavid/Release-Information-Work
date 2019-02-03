# Release-Information-Work
Using Selenium, Openpyxl and Google Sheet Queries to transform release information at work

**Problem:** At work collecting information about past releases was difficult. Information was stored in several locations and the most centralized, a Sharepoint site maintained by the client, had every test behind panels that were loaded by JavaScript after the panel was opened.

**Solution:** Scrape the Sharepoint site using Python and Selenium and write the results to an Excel workbook using Openpyxl. The results were then moved to a Google Sheet where queries could be run to spot trends and better analyze the data. 

**Next Steps: The script is still in use and used to collect the most recent set of tests. 
To further improve the process:  
- The script could be run automatically to detect new releases on the Sharepoint site
- The results could be written directly to the Google Sheet and not transfered from Excel manually.


**Sample Google Sheet:** https://drive.google.com/file/d/1PbgWVXZxY_ZazGD0Ef3toYouDnERvFcO/view?usp=sharing
Image of Sharepoint Site

