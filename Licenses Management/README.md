
# License-Users

### Assign Licenses to users, using a CSV input, and using another user as template

For admins who need to assign Office365 licenses and do not have an Azure P1 plan (Cannot use the group license assignment feature).

This script will let you input a CSV file containig a list of users and relative template user for licensing mirroring.

#### Usage:
* Configure a user with the required licenses, like E5 Plan and activate desiders workloads, like Exchange,SharePoint and Office(or all, not limiting)

* Create a CSV file with this format<br>
    EmailAddress | TemplateEmail | UsageLocation [Optional]
    ------------ | ------------- | -------------  
    Content from cell 1 | Content from cell 2 | 2 letters Country Code [ Optional]

    ##### CSV <br>
    `EmailAddress,TemplateEmail,UsageLocation` <br>
    `user@contoso.com,Template@contoso.com,US` <br>

* Import the function in PowerShell using `. .\License-Users.ps1`

* Execute the import in any of these ways:
    * `License-Users -CSVFile c:\temp\users.csv`
    * `License-Users -CSVFile c:\temp\users.csv -UsageLocation AU`
    * `License-Users -CSVFile c:\temp\users.csv -LogDirectory c:\temp`
    * `License-Users -CSVFile c:\temp\users.csv -LogDirectory c:\temp -UsageLocation AU`
 <br>
 <br>


##### -CSVFile [path]\file.csv<br> 
The only "required" parameter is the -CSVFile, but check UsageLocation


##### -UsageLocation AU<br> 
It is the UsageLocation parameter in Get-ADUser or Get-MsolUser.
UsageLocation is a must have for a user to be able to receive a license, so the script will check if the AD User in Office365 already has it.
You have multiple options here:

* Add the option column UsageLocation to the CSV file, 
    * if the user has no location already, the one in the column will be used
* Pass the parameter to the script
    * the passed value will be used for all users in CSV, where value is missing on the user in the Cloud 
* Do not pass it at all
    * Script will check anyway for each user if it is present, if it will be the case, then no action will be required, else you will prompted to enter the country code and asked if you want to use the provided one as the default for all the users missing it. <br> 
    If you refuse to use it as default, you will get a prompt for each one missing.

##### -LogDirectory <br>
Script will try to create the directory if non existent, if provided a log with the following name format will be created upon execution `YYYY-MM-DD_HH-MM_Licenses.log` ---> `2018-08-13_15-39_Licenses.log`

