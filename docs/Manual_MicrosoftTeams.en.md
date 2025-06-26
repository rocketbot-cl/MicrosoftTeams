



# MicrosoftTeams
  
This module allows you to connect to the Microsoft Teams API to create and manage teams, groups, and channels  

*Read this in other languages: [English](Manual_MicrosoftTeams.md), [Português](Manual_MicrosoftTeams.pr.md), [Español](Manual_MicrosoftTeams.es.md)*
  
![banner](imgs/Banner_MicrosoftTeams.png o jpg)
## How to install this module
  
To install the module in Rocketbot Studio, it can be done in two ways:
1. Manual: __Download__ the .zip file and unzip it in the modules folder. The folder name must be the same as the module and inside it must have the following files and folders: \__init__.py, package.json, docs, example and libs. If you have the application open, refresh your browser to be able to use the new module.
2. Automatic: When entering Rocketbot Studio on the right margin you will find the **Addons** section, select **Install Mods**, search for the desired module and press install.  

## How to use this module

Before using this module, you need a corporate email account and must register your application in the Azure App Registrations portal.

1. Sign in to the Azure portal (Applications Registration: https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade ).
2. Select "New registration".
3. Under “Supported account types” supported choose:
    - "Accounts in any organizational directory (any Azure AD directory: multi-tenant) and personal Microsoft accounts (such as Skype or Xbox)" for this case use Tenant ID = **common**.
    - "Only accounts from this organizational directory (only this account: single tenant) for this case use application-specific **Tenant ID**.
    - "Personal Microsoft accounts only" for this case use use Tenant ID = **consumers**.
4. Set the redirect uri (Web) as: https://localhost:5001/ and click "Register".
5. Copy the application (client) ID. You will need this value.
6. Under "Certificates and secrets", generate 
a new client secret. Set the expiration (preferably 24 months). Copy the **VALUE** of the created client secret (NOT the Secret ID). It will hide after a few minutes.
7. Under "API permissions", click "Add a permission", select "Microsoft Graph", then "Delegated permissions", find and select **"Channel.ReadBasic.All"**, **"ChannelMessage.Read.All"**, **"ChannelMessage.ReadWrite"**, **"ChannelMessage.Send"**, **"Chat.Create"**, **"Chat.ManageDeletion.All"**, **"Chat.Read"**, **"Chat.ReadWrite"**, **"Chat.ReadWrite.All"**, **"Team.Create"**, **"Team.ReadBasic.All"**, **"TeamMember.Read.All"**, **"TeamMember.ReadWrite.All"** and finally "Add permissions".

8. Access code, generate code by entering the following link:

https://login.microsoftonline.com/{tenant}/oauth2/v2.0/authorize?client_id={**client_id**}&response_type=code&redirect_uri={**redirect_uri**}&response_mode=query&scope=offline_access%20Channel.ReadBasic.All%20ChannelMessage.Read.All%20ChannelMessage.ReadWrite%20ChannelMessage.Send%20Chat.Create%20Chat.ManageDeletion.All%20Chat.Read%20Chat.ReadWrite%20Chat.ReadWrite.All%20Team.Create%20Team.ReadBasic.All%20TeamMember.Read.All%20TeamMember.ReadWrite.All&state=12345
Replace within the link {tennat}, {client_id} and {redirect_uri}, with the data corresponding to the created application.

9. If the operation was successful, the browser URL will change to: http://localhost:5001/?code={**CODE**}&state=12345#!/
The value that appears in {CODE}, copy it and use it in the Rocketbot command in the "code" field to make the connection.

Note: The browser will NOT load any pages.

## Description of the commands

### Set credentials
  
Set credentials to make available the API
|Parameters|Description|example|
| --- | --- | --- |
|client_id|Client ID obtained in the creation of the application|Your client_id|
|client_secret|Client secret obtained in the creation of the application|Your client_secret|
|redirect_uri|Application redirect URL|http://localhost:5000|
|code|Data obtained by placing the authentication URL. Check the documentation for more information|code|
|tenant|Tenant identifier to which you want to connect|tenant|
|Result|Variable to store result. If the connection is successful, it will return True, otherwise it will return False|connection|
|session|Variable to store the session identifier. Use in case you want to connect to more than one account at the same time|session|

### Create Team
  
Create a new team
|Parameters|Description|example|
| --- | --- | --- |
|Name|Team name|Rocketbot|
|Description|Description of Team (optional)|Team Rocketbot|
|Visibility|Visibility Public or Private|Public|
|Result|Variable to store result. If the task is successful, it will return True, otherwise it will return False|res|
|session|Session ID|session|

### List Teams
  
list Joined Teams
|Parameters|Description|example|
| --- | --- | --- |
|Result|Variable where the query result will be saved|res|
|session|Session ID|session|

### Get Team Details
  
Get details of a specific team
|Parameters|Description|example|
| --- | --- | --- |
|Identificador de equipo|Team ID|Team ID|
|Result|Variable name where the result will be saved|res|
|session|Session ID|session|

### Delete Team
  
Delete a specific team
|Parameters|Description|example|
| --- | --- | --- |
|Identificador de equipo|Team ID|ID Team|
|Result|Variable name where the result will be saved|res|
|session|Session ID|session|

### List Members
  
List Members of a specific team
|Parameters|Description|example|
| --- | --- | --- |
|Identificador de equipo|Team ID|ID Team|
|Result|Variable name where the result will be saved|res|
|session|Session ID|session|

### Add Member
  
Add Member of a specific team
|Parameters|Description|example|
| --- | --- | --- |
|Team|Team ID|ID Team|
|User ID|User ID|ID|
|User's email address|User's email address|test@test.com|
|Result|Variable name where the result will be saved|res|
|session|Session ID|session|

### Remove Member
  
Remove member of a specific team
|Parameters|Description|example|
| --- | --- | --- |
|Team|Team ID|ID Team|
|User ID|User ID|User ID|
|Result|Variable name where the result will be saved|res|
|session|Session ID|session|

### list Channels
  
List channels within a team
|Parameters|Description|example|
| --- | --- | --- |
|Team ID|Team ID to list channels|Team ID|
|Order by|Parameters to order the results of the query made|name desc|
|Filtrar por|Filter to apply to perform the query|name eq 'file.txt'|
|Quantity|Number of items to obtain. It will return the top items of the query|10|
|Result|Variable name to save the result|res|
|session|Session ID|session|

### Get Channel Details
  
Get details of a specific channel
|Parameters|Description|example|
| --- | --- | --- |
|Team ID|Team ID|Team ID|
|Channel ID|Channel ID.|Channel ID|
|Result|Variable name to save the result|res|
|session|Session ID|session|

### Create Channel
  
Create a new channel in a team
|Parameters|Description|example|
| --- | --- | --- |
|Team ID|Team ID|23XWM5ASR67M67S6KYNCV66KFMQFOTOPDL|
|Name||Name|
|Description (Optional)||Description (Optional)|
|Result|Variable to store result. If the task is successful, it will return True, otherwise it will return False|res|
|session|Session ID|session|

### Delete Channel
  
Delete a channel from a team
|Parameters|Description|example|
| --- | --- | --- |
|Team ID|Team ID.|Team ID|
|Channel ID|Channel ID.|channel id|
|Result|Variable to store result. If the task is successful, it will return True, otherwise it will return False|res|
|session|Session ID|session|

### list Messages
  
List messages in a channel
|Parameters|Description|example|
| --- | --- | --- |
|Team ID||Team ID|
|Channel ID||15ZLM4OKQTAC3M7UDDR5DBUKPA4U8ULNXW|
|Result|Variable to store result. If the task is successful, it will return True, otherwise it will return False|res|
|session||session|

### Get Message Details
  
Get details of a specific message in a channel
|Parameters|Description|example|
| --- | --- | --- |
|Team ID|Team ID|id|
|Channel ID|Channel ID|id|
|Message ID|Message ID|id|
|Result|Variable to store result. If the task is successful, it will return True, otherwise it will return False|res|
|session|Session ID|session|

### Send Channel Message
  
Send a message to a channel
|Parameters|Description|example|
| --- | --- | --- |
|Team ID|Team ID|id|
|Channel ID|Channel ID|id|
|Message body|Message body|content|
|Subject|Subject|subject|
|Result|Variable to store result. If the task is successful, it will return True, otherwise it will return False|res|
|session|Session ID|session|
