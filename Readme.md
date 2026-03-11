# python_acas_report_automation
ACAS Report Automation by python script

This script was created to automate the collection and organization of reports from Security Center / ACAS after scans are completed.

Instead of manually logging into Security Center, finding the latest report results, downloading each file one by one, and organizing them into monthly folders, the script performs the entire process automatically.

Its main purpose is to save time, reduce repetitive manual work, and ensure that important security reports are consistently collected and stored in a structured location.

It downloads the most recent available report results, including the vulnerability CSV report (I had to create report for this CSV as vulnerability detail) and key PDF reports such as the Critical and Exploitable Vulnerabilities Report, Monthly Executive Report, and Remediation Instructions by Host Report.

The script also improves usability by automatically converting the downloaded CSV report into an Excel file, making the data easier to review and work with. In addition, it organizes all downloaded files into year and month folders so reports can be archived in a clean and predictable way.

Overall, this file was made to support a more efficient reporting workflow in an Security Center / ACAS environment, especially in situations where reports need to be collected regularly and stored for tracking, review, or compliance purposes.
