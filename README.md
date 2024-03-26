# Office Deployment Repository
Contains a few xml files for office deployment using Office Deployment Tool (ODT). This repository also contains a script that will uninstall all office versions and install the selected version of your choice.
**NOTE: This tool does not support the MSI version of Office 2016**

# How To Use:
You can run the script in multiple ways. You can run it through PowerShell Core (pwsh), PowerShell, Command Prompt (cmd), or just click on the .ps1 file. If you are not using the script, you can download the appropriate .xml file, download the [ODT](https://go.microsoft.com/fwlink/p/?LinkID=626065), extract the odt (just run the file), and run setup.exe /download {myxmlfile.xml} (without the brackets), then run setup.exe /configure {myxmlfile.xml}.

If you want to quickly run the script without directly downloading the file, then this command will do the trick:

Invoke-WebRequest -UseBasicParsing "https://www.github.com/Cryozyme/office-deployment/raw/main/office_deployment.ps1" | Invoke-Expression