# MYOB-Windows-Service Sample Application

Windows Service developed in c# to for reporting and automatic backup of sales and purchase orders of company file in myob cloud.


## C# Windows Service
C# Sample app demonstrating accessing the AccountRight Live API using the SDK in c# and creating windows serivce for periodic automatic backup of files in excel format.

### AccountRight API - C#
*	Uses the new .Net SDK  
*	Navigates available files on cloud server 
*	Manages OAuth login (note: SDK manages refreshing of tokens!) 
*	Demonstrates use of paging and filtering through listing of Purchase order Sales, Profit and loss, and balance sheets.
*	Gives automatic backup of purchase order and sales report of definite period of time which we can configure in App.config.

### Requirements
*	Visual Studio 
*	MYOBApi developer key, secret key, and Code (To access cloud server)

### Getting up and running
* Unzip source code to a local folder and open the solution file. 
* Restore Nuget Packages,
* Get Developer key and Developer Secret key for your account from developer section from account. For more detail go to 
[this page](http://developer.myob.com/api/accountright/api-overview/getting-started/)
* Update App.cofnig in MYOBCustomService Project's App config.
* Rebuild Solution and run the installation file in Debug/Relase folder of MyOBInstaller Project. 
* Give username and password of Windows account. 
* Go to services from run using services.msc . Run MyOBCustomeService. 
* All done. File will be created in location which you have set in in App.config.

**Happy Coding!!!**
