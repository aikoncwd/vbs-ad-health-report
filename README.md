# Description
VBS Script that check's your Active Directory Health status and report using e-mail. It work's for complex Domains structures, checking every DC for each Site. It also checks basic hardware information (like Disk usage and free RAM). There are few options to customize how and when you want the reports.

![](https://i.imgur.com/VLrYsET.png)
![](https://i.imgur.com/Qv8kmnQ.png)

# How to install / use
Download a raw copy of the `vbs-ad-health-report.vbs` script, save it in your main Active Directory Domain Controller (for example C:\scripts\vbs-ad-health-report.vbs). Open the script with any text editor ([Notepad++](https://notepad-plus-plus.org/download) recommended).

![](https://i.imgur.com/lyp9vBh.png)

**usingOU** is an important variable. If you set it *True*, the script will query your Domain and list any server stored in the default OU for DC server. If you have your DC servers stored in a custom OU, you can specify in the **organizationUnitDC** variable. If you are under a Domain with complex structure I recomend to set **usingOU** = *False*, then specify the name of every DC server that you want to monitor and report.

The other variables (options) are self explanatory, you enable/disable the hardware report with **hardwareReport** variable. Othe lines are for customize the levels of reports (e-mail, CSV attachment, outputFile, ...), just change **True** or **False**. Configure then your e-mail server settings (*GMail* supported!). Save changes and double-click to execute and test the script. If everyting works fine, create a new scheduled task to run this script automatically (1 run per day recommended):

![](https://i.imgur.com/JgS151V.png)

**Remember to run this scheduled-task as Administrator AND with elevated privileges**

You will get a daily report of your Active Directory Health Status in your e-mail, like [this](https://i.imgur.com/X3TfbYT.png)

# Credits
Original idea by [Vikas Sukhija](https://gallery.technet.microsoft.com/scriptcenter/Active-Directory-Health-709336cd), I used his idea, translated into VBS script and adding more AD checks/tests.
