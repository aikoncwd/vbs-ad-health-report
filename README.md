# vbs-ad-health-report
VBS Script that check's your Active Directory Health status and report using e-mail. It work's for complex Domains structures, checking every DC for each Site.

![](https://i.imgur.com/zyd7YiY.png)

## How to install / use
Download a raw copy of the `vbs-ad-health-report.vbs` script, save it in your main Active Directory Domain Controller (for example C:\scripts\vbs-ad-health-report.vbs). Open the script with any text editor ([Notepad++](https://notepad-plus-plus.org/download) recommended). At **line 130** you will find a function called `SendMail()`, edit and configure your own e-mail settings (*GMail* supported!). Save changes and double-click to execute and test the script.

If everyting works fine, create a new scheduled task to run this script automatically (1 run per day recommended):

![](https://i.imgur.com/JgS151V.png)

You will get a daily report of your Active Directory Health Status in your e-mail, like [this](https://i.imgur.com/X3TfbYT.png)

## Credits
Original idea by [Vikas Sukhija](https://gallery.technet.microsoft.com/scriptcenter/Active-Directory-Health-709336cd), I used his idea, translated into VBS script and adding more AD checks/tests.
