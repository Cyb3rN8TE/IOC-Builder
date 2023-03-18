# IOC Builder

![alt text](https://github.com/nwjohns101/IOC-Builder/blob/Dev/Images/logo.png) 

**PowerShell script that refangs or defangs IP addresses, domains and URLs in a spreadsheet (.CSV) containing Indicators of Compromise (IOC)**

**Why did I make this?**

Indicators of Compromise (IOC) are often provided with defanged values in spreadsheet format. My tool can sort through the data, refanging any IP addresses which makes it easier to import into other programs etc.

**What does it do?**
1) Imports raw .CSV that the user selects
2) Searches the spreadsheet for all IP addresses, domains and URLs
3) Refangs or defangs any values that are IP addresses, domains or URLs
5) Saves the resulting .CSV file with a name and location of the users selection. 

**Prerequisites:**

- Windows 7 with Service Pack 1 (SP1)
- Windows Vista with Service Pack 2 (SP2)
- Windows 8
- Windows 10
- Windows 11

**How to run the tool:**
1) Right click iocbuilder.ps1 and select 'run with PowerShell'
