# OuputAlertsRulesMonitorsToCSV PowerShell Script

## [Download Here][Download]

[Download]: https://github.com/thekevinholman/OuputAlertsRulesMonitorsToCSV/archive/master.zip

### SCOM - Get All SCOM Rules and Monitors with their Alert details to a CSV

https://kevinholman.com/2018/08/11/get-all-scom-rules-and-monitors-with-their-alert-details-to-a-csv/

Version History:
* 1.6  (04-20-2022)
	* Added Alert information for rules using System.Health.GenerateAlertForType
* 1.5  (09-02-2021)
	* Combined rules and monitors into a single CSV output
	* Added Target class DisplayName and Name
	* Added field WorkFlowType for Rule or Monitor 	
* 1.4  (09-02-2021)
	* Fixed bug when Alert did not have an AlertMessageId string defined
	* Added command line params for -OutputDir and -ManagementServer
