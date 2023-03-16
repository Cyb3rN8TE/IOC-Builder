# IOC Builder

![alt text](https://github.com/nwjohns101/IOC-Builder/blob/Dev/Images/logo.png) 

**PowerShell script that refangs and organises IP addresses, domains and URLs in a spreadsheet (.XLSX or .CSV) containing Indicators of Compromise (IOC)**

**Why did I make this?**

Indicators of Compromise (IOC) are often uncategorised when provided in spreadsheet format. My tool sorts through the data, refangs any IP addresses and makes it easier to copy and paste the data elsewhere.

**What does it do?**
1) Imports raw .XLSX or .CSV that the user selects
2) Searches the spreadsheet for all IP addresses, domains and URLs
3) Refangs any values that are IP addresses, domains or URLs
4) Places these values in a new columns under the appropriate heading.
5) Saves the resulting .XLSX or .CSV file with a name and location of the users selection. 

**Prerequisites:**

Microsoft Office 2016 installed on system

Microsoft Windows OS:
- Windows 7 with Service Pack 1 (SP1)
- Windows Vista with Service Pack 2 (SP2)
- Windows 8
- Windows 10
- Windows 11

**Dependencies:**
- Microsoft.Office.Interop.Excel.dll in the root directory of the script (provided)

**How to run the tool:**
1) Right click LRWC Log Beautify.ps1 and select 'run with PowerShell'
