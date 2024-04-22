# M365 Copilot Audit Power BI Report
This report is using data from Purview Audit log and Entra ID exports into two CSV files. The first csv file is storing Copilot Interaction events and other csv is exporting user details (Display Name, UPN, Position, City, Country) for users that have M365 copilot license assigned.
Power BI report is reading data from these two files that are stored on SharePoint Online Document Library. To sync exported files to SharePoint, you need to create (or use existing) SharePoint Site (Communication or Teams) and sync Document Library to Windows PC where PowerShell scripts will get executed.

![screenshot](/img/Brand.png)

## Prerequisites
1.	M365 Copilot Licenses assigned to users and some history with it :)
2.	Power BI Desktop (most recent version) [Download link] (https://aka.ms/pbidesktopstore)
3.	Audit Log Read Permissions (Security Reader is minimum)
4.	Windows PowerShell 7 [Download link] (https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.4)
## Initial configuration
1.	Open **Audit-Get-Events.ps1** and edit the folderPath (line 27).
2.	Open **Audit-Get-Users.ps1** and edit the folderPath (line 29).
3.	Open **Copilot_TimeStamp.txt** and configure start date for your report data.
## Extracting AD users
1.	Open **Windows PowerShell (not PS7)** as Administrator
2.	Run **Audit-Get-Users.ps1** from Windows PowerShell
3.	Once you get prompted to authenticate, authenticate with an account that has at least Security Reader Permissions (Global Admin will work of course).
4.	Once the script is complete, the folder will include **Copilot_Users.csv** that will contain list of users from your tenant that have M365 Copilot License assigned.
## Extracting Copilot Interactions
1.	Run PowerShell 7 as Administrator
2.	Run **Audit-Get-Events.ps1** from Windows PowerShell
3.	Once you get prompted to authenticate, authenticate with an account that has at least Security Reader Permissions (Global Admin will work of course).
4.	Once the script is complete, the folder will include **Copilot_Events.csv** that will contain list of all events from November 1st, 2023, until today.
## Power BI Template configuration
1. Open M365 Copilot Audit Report.pbit with Power BI Desktop
2. You will get prompted with M365 Copilot Audit Report configuration screen

![screenshot](/img/Picture1a.png)

3. Populate it with parameters captured in Initial configuration section.
4. Press Load button and click Connect
5. Report will start to pull data from two files

![Alt text](/img/Picture3.png?raw=true)

6. Once the process is complete, your report will display information

![Alt text](/img/Picture4.png?raw=true)

7. Save the Report as **M365 Copilot Audit Report.pbix** to your PC
## Anonymizing users
In case you want to anonymize user Display names in Power BI report:

![Alt text](/img/Picture5.png?raw=true)

1. Click on the Employee visual
2. Click x next to Employee in Visualizations pane to remove employees.
3. Tick the box on Anonymous in Data Pane

![Alt text](/img/Picture6.png?raw=true)

4. Repeat the same steps for Top Active users visual
