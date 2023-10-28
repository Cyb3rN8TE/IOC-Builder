# IOC Builder

![alt text](https://github.com/Cyb3rN8TE/IOC-Builder/blob/Dev/Images/Logo.png) 

**PowerShell script that refangs or defangs IP addresses, domains and URLs in a spreadsheet (.CSV) containing Indicators of Compromise (IOC)**

**Why did I make this?**

Indicators of Compromise (IOC) are often provided with defanged values in spreadsheet format. My tool can sort through the data, refanging any IP addresses which makes it easier to import into other programs etc.

**What does it do?**
1) Imports raw .CSV that the user selects
2) Searches the spreadsheet for all IP addresses, domains and URLs
3) Refangs or defangs any values that are IP addresses, domains or URLs
5) Saves the resulting .CSV file with a name and location of the users selection. 
