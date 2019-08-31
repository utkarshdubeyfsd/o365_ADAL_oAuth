# O365 ADAl Oauth

It is a console application built on .NET Framework. Which will first do authentication via Graph API and then accuire a Token. 

This is the list of operation which you can perform using this console application.

#### O365 User Operation

- Retrieve User List

#### Outlook Operation

- Create Calendar Events
- Edit Existing Calendar Events
- Delete Calendar Events
- Retrieve Calendar Events

#### SharePoint Operation

- Upload a File
- Download a File
- Copy File(between libraries)
- Delete File
- Get All File Names and Folder Names
- Create Folder

## Register an app on Azure

Register an app using [Register an application with the Microsoft identity platform](https://docs.microsoft.com/en-us/graph/auth-register-app-v2) article. After that, you need to acquire Client ID, Client Secret, and Tenant ID.

## Set permission

You will need to set the following permission in your Azure app.

![Permission](https://github.com/utkarshdubeyfsd/o365_ADAL_oAuth/blob/master/Permission.PNG)

> After setting up an application and applying for permission, In my case, it took 1 day for the application to run. Otherwise, it shows the Token failed to acquire.

## Using this application

To build and start using this application, follow below mentioned instructions.

1. Clone this repository by executing the following command in your console:

```
git clone https://github.com/utkarshdubeyfsd/o365_ADAL_oAuth.git
```

2. Install all the packages from NuGet
   - Microsoft.Extension.Configuration
   - Microsoft.Extension.Configuration.Binder
   - Microsoft.Extension.Configuration.Json
   - Microsoft.Graph
   - Microsoft.Identity.Client
   - Microsoft.NETCore.App
   - Newtonsoft.Json
   - Syroot.Windows.IO.KnownFolders
3. Update details on **appsettings.json** file like `ClientId`, `ClientSecret` and etc.
4. Execute the application

## Output

This is how the application will look like.

![Output](https://github.com/utkarshdubeyfsd/o365_ADAL_oAuth/blob/master/o365_ADAL_Oauth_output.gif)


> No Wor**d**, Only Wor**k**
