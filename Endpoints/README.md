# EndPoints to Csv

### Extract Office 365 endpoints IP Addresses and generate csv file from it

  Get the Office 365 Endpoints using the webservice endpoint and generate per service csv file (and per service\tcp port txt files). <br>

  Once the CSV files are generated, open them with Excel and convert the content to a table + use the Top Align feature after selecting all the content.<br>
  
  To convert only the csv content, and not all the csv to a table in Excel, try:
  * Open the file with Excel
      * Select cell A:1
      * Keep ctrl + Shift pressed
         * press Right Arrow 
         * press Down Arrow
      * Use the convert to table from the menu` and then select the who

#### Usage <br>
Use the script in any of the following way <br>

    * `.\EndPointsToCsv.ps1`
    * `.\EndPointsToCsv.ps1 -Path "c:\temp\"`
    * `.\EndPointsToCsv.ps1 -Path "c:\temp\" -GenerateTXT`
    * `.\EndPointsToCsv.ps1 -GenerateTXT`



#### Parameters <br>

###### $Path <br>
"path\" where to save the csv files. If omitted ".\" will be used.

###### $GenerateTXT <br>
Switch parameter, if passed with -GenerateTXT it will cause the scrip to generate a per Service folder, containing a list of txt files named as the TCP port used