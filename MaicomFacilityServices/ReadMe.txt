Deployment instructions

This script is too big to be run through MSSMS.
This script needs to be run throgh the command line using the SQLCMD utility.

- Open up and command prompt window
- Cd to the folder where the SQL is saved then
- type the following command: sqlcmd -i 01-WO84109-Maicom_BulkDocInsert.sql -o Results.txt [Enter]

Dpending of the load and performance of the server the script will take more than one hour to complete
Look at the Results.txt file for any errors