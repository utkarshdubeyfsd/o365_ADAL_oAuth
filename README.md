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

## Steps which you will need to perform before running the application

1. Clone the application to your local environment.
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
