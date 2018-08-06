# san-zoning
Small script on powershell to automate process of zoning fibre channel

Require Powershell 5 or higher

* Connect to your initiators, save list of wwn to excel file Initiator.xlsx
* Connect to your targets, save list of wwns to excel file Targets.xlsx
* Fill excel file Zoning.xlsx, add Initiator <--> Target links
* run powershell main.ps1, save output to file
* paste output to san-switch ssh session

