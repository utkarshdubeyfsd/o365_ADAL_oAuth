using Microsoft.Graph;
using Microsoft.Identity.Client;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Security.Cryptography.X509Certificates;
using System.Threading;
using System.Threading.Tasks;

namespace O365_ADAL_OAuth
{
    /// <summary>
    /// This sample shows how to query the Microsoft Graph from a daemon application
    /// which uses application permissions.
    /// For more information see https://aka.ms/msal-net-client-credentials
    /// </summary>
    class Program
    {

        public static string option = @"";

        /// <summary>
        /// Main Method
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            try
            {
                RunAsync().GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }

            Console.WriteLine("Press any key to exit");
            Console.ReadKey();
        }

        /// <summary>
        /// This method will run each job asynchronously
        /// </summary>
        /// <returns></returns>
        private static async Task RunAsync()
        {

            AuthenticationConfig config = AuthenticationConfig.ReadFromJsonFile("appsettings.json");

            // Even if this is a console application here, a daemon application is a confidential client application
            IConfidentialClientApplication app;

#if !VariationWithCertificateCredentials
            app = ConfidentialClientApplicationBuilder.Create(config.ClientId)
                .WithClientSecret(config.ClientSecret)
                .WithAuthority(new Uri(config.Authority))
                .Build();
#else
            X509Certificate2 certificate = ReadCertificate(config.CertificateName);
            app = ConfidentialClientApplicationBuilder.Create(config.ClientId)
                .WithCertificate(certificate)
                .WithAuthority(new Uri(config.Authority))
                .Build();
#endif

            // With client credentials flows the scopes is ALWAYS of the shape "resource/.default", as the 
            // application permissions need to be set statically (in the portal or by PowerShell), and then granted by
            // a tenant administrator
            string[] scopes = new string[] { "https://graph.microsoft.com/.default" };

            AuthenticationResult result = null;

            // Acquiring a token to call the API
            try
            {
                result = await app.AcquireTokenForClient(scopes)
                    .ExecuteAsync();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Token acquired");
                Console.ResetColor();
            }
            catch (MsalServiceException ex) when (ex.Message.Contains("AADSTS70011"))
            {
                // Invalid scope. The scope has to be of the form "https://resourceurl/.default"
                // Mitigation: change the scope to be as expected
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Scope provided is not supported");
                Console.ResetColor();
            }

            if (result != null)
            {
                var httpClient = new HttpClient();
                var apiCaller = new ProtectedApiCallHelper(httpClient);
                var spApiCaller = new SPApiCallHelper(httpClient);

                #region Calling APIs
                while (true)
                {
                    Console.WriteLine("--------------------");
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.Write(new string('\n', 2));
                    Console.WriteLine("o Choose Option");
                    Console.WriteLine();
                    option = @"
1. Retrieve User List
2. Calendar Events
3. Files/Folders
4. Exit
";
                    Console.WriteLine(option);
                    Console.ResetColor();
                    Console.Write("Option: ");
                    switch (Console.ReadLine())
                    {
                        case "1": /*RETRIEVE USER LIST*/
                            Console.WriteLine();
                            //retrieve user list from o365
                            await apiCaller.CallWebApiAndProcessResultASync("https://graph.microsoft.com/v1.0/users", result.AccessToken, Display);
                            Console.WriteLine();
                            break;
                        case "2": /*CALENDAR EVENTS*/
                            Console.WriteLine("--------------------");
                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.Write(new string('\n', 2));
                            Console.WriteLine("o Choose Option");
                            Console.WriteLine();
                            option = @"
1. Create Calendar Events
2. Edit Existing Calendar Events
3. Delete Calendar Events
4. Retrieve Calendar Events
5. Exit
";
                            Console.WriteLine(option);
                            Console.ResetColor();
                            Console.Write("Option: ");
                            switch (Console.ReadLine())
                            {
                                case "1": /*CREATE CALENDAR EVENT*/
                                    string eventBody = "", subject = "", location = "", start = "", end = "";
                                    List<Attendee> listAttendee = new List<Attendee>();
                                    CultureInfo provider = CultureInfo.InvariantCulture;

                                    Console.WriteLine();
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    Console.WriteLine("Fill following details: ");
                                    Console.Write("Subject: ");
                                    subject = Console.ReadLine();
                                    Console.Write("Start Date(dd/MM/yyyy HH:mm:ss): ");
                                    start = (DateTime.ParseExact(Console.ReadLine(), "dd/MM/yyyy HH:mm:ss", provider).ToUniversalTime()).ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
                                    Console.Write("End Date(dd/MM/yyyy HH:mm:ss): ");
                                    end = (DateTime.ParseExact(Console.ReadLine(), "dd/MM/yyyy HH:mm:ss", provider).ToUniversalTime()).ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
                                    Console.Write("Description: ");
                                    eventBody = Console.ReadLine();
                                    Console.Write("Location: ");
                                    location = Console.ReadLine();
                                    Console.Write("No. of Attendees: ");
                                    Int32 num = Convert.ToInt32(Console.ReadLine());
                                    for (var i = 0; i < num; i++)
                                    {
                                        string emailaddress = "adelev@contoso.onmicrosoft.com", name = "Adele Vance";
                                        Console.WriteLine("Attendee " + i);
                                        Console.Write("Attendee name: ");
                                        name = Console.ReadLine();
                                        Console.Write("Attendee email address: ");
                                        emailaddress = Console.ReadLine();
                                        listAttendee.Add(new Attendee
                                        {
                                            EmailAddress = new EmailAddress
                                            {
                                                Address = emailaddress,
                                                Name = name
                                            },
                                            Type = AttendeeType.Required
                                        });
                                    }
                                    Console.WriteLine();
                                    Console.ResetColor();

                                    var @event = new Event
                                    {
                                        Subject = subject,
                                        Body = new ItemBody
                                        {
                                            ContentType = BodyType.Text,
                                            Content = eventBody
                                        },
                                        Start = new DateTimeTimeZone
                                        {
                                            DateTime = start,
                                            TimeZone = "UTC"
                                        },
                                        End = new DateTimeTimeZone
                                        {
                                            DateTime = end,
                                            TimeZone = "UTC"
                                        },
                                        Location = new Location
                                        {
                                            DisplayName = location
                                        },
                                        Attendees = listAttendee
                                    };

                                    JObject body = (JObject)JToken.FromObject(@event);
                                    //create calendar events in o365
                                    await apiCaller.post("https://graph.microsoft.com/v1.0/users/" + config.userPrincipalName + "/events", result.AccessToken, body, Display);
                                    break;
                                case "2": /*EDIT CALENDAR EVENTS*/
                                    string editSub = "";

                                    Console.WriteLine();
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    Console.Write("Enter subject of Calendar: ");
                                    editSub = Console.ReadLine();
                                    Console.ResetColor();
                                    Console.WriteLine();

                                    if (String.IsNullOrEmpty(editSub))
                                    {
                                        Console.WriteLine("Invalid Subject");
                                    }
                                    else
                                    {
                                        string edit_eventBody = "", edit_subject = "", edit_location = "", edit_start = "", edit_end = "";
                                        List<Attendee> edit_listAttendee = new List<Attendee>();
                                        CultureInfo edit_provider = CultureInfo.InvariantCulture;

                                        Console.ForegroundColor = ConsoleColor.Yellow;
                                        Console.WriteLine("Fill following details: ");
                                        Console.Write("Subject: ");
                                        edit_subject = Console.ReadLine();
                                        Console.Write("Start Date(dd/MM/yyyy HH:mm:ss): ");
                                        edit_start = (DateTime.ParseExact(Console.ReadLine(), "dd/MM/yyyy HH:mm:ss", edit_provider).ToUniversalTime()).ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
                                        Console.Write("End Date(dd/MM/yyyy HH:mm:ss): ");
                                        edit_end = (DateTime.ParseExact(Console.ReadLine(), "dd/MM/yyyy HH:mm:ss", edit_provider).ToUniversalTime()).ToString("yyyy-MM-ddTHH:mm:ss.fffZ");
                                        Console.Write("Description: ");
                                        edit_eventBody = Console.ReadLine();
                                        Console.Write("Location: ");
                                        edit_location = Console.ReadLine();
                                        Console.Write("No. of Attendees: ");
                                        Int32 edit_num = Convert.ToInt32(Console.ReadLine());
                                        for (var i = 0; i < edit_num; i++)
                                        {
                                            string emailaddress = "adelev@contoso.onmicrosoft.com", name = "Adele Vance";
                                            Console.WriteLine("Attendee " + i);
                                            Console.Write("Attendee name: ");
                                            name = Console.ReadLine();
                                            Console.Write("Attendee email address: ");
                                            emailaddress = Console.ReadLine();
                                            edit_listAttendee.Add(new Attendee
                                            {
                                                EmailAddress = new EmailAddress
                                                {
                                                    Address = emailaddress,
                                                    Name = name
                                                },
                                                Type = AttendeeType.Required
                                            });
                                        }
                                        Console.WriteLine();
                                        Console.ResetColor();

                                        var @edit_event = new Event
                                        {
                                            Subject = edit_subject,
                                            Body = new ItemBody
                                            {
                                                ContentType = BodyType.Text,
                                                Content = edit_eventBody
                                            },
                                            Start = new DateTimeTimeZone
                                            {
                                                DateTime = edit_start,
                                                TimeZone = "UTC"
                                            },
                                            End = new DateTimeTimeZone
                                            {
                                                DateTime = edit_end,
                                                TimeZone = "UTC"
                                            },
                                            Location = new Location
                                            {
                                                DisplayName = edit_location
                                            },
                                            Attendees = edit_listAttendee
                                        };

                                        JObject edit_body = (JObject)JToken.FromObject(@edit_event);

                                        //update calendar events in o365
                                        await apiCaller.put("https://graph.microsoft.com/v1.0/users/" + config.userPrincipalName + "/events?$select=id&$filter=subject eq '" + editSub + "'", result.AccessToken, edit_body, Display);
                                    }
                                    break;
                                case "3": /*DELETE CALENDAR EVENTS*/
                                    string delSubject = "";

                                    Console.WriteLine();
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    Console.Write("Enter subject of Calendar: ");
                                    delSubject = Console.ReadLine();
                                    Console.WriteLine();
                                    Console.ResetColor();

                                    if (delSubject == "")
                                    {
                                        //retrieve calendar event where subject is specific, from o365
                                        await apiCaller.delete("https://graph.microsoft.com/v1.0/users/" + config.userPrincipalName + @"/events?$select=id&$filter=subject eq null", result.AccessToken, Display);
                                    }
                                    else
                                    {
                                        //retrieve calendar event where subject is specific, from o365
                                        await apiCaller.delete("https://graph.microsoft.com/v1.0/users/" + config.userPrincipalName + @"/events?$select=id&$filter=subject eq '" + delSubject + "'", result.AccessToken, Display);
                                    }

                                    break;
                                case "4": /*GET CALENDAR EVENTS*/
                                    Console.WriteLine();
                                    //retrieve calendar events from o365
                                    await apiCaller.CallWebApiAndProcessResultASync("https://graph.microsoft.com/v1.0/users/" + config.userPrincipalName + @"/events?$select=subject,body,attendees,start,end,location", result.AccessToken, Display);
                                    Console.WriteLine();
                                    break;
                                case "5": /*EXIT*/
                                    Environment.Exit(0);
                                    break;
                                default: /*INVALID OPTION*/
                                    Console.WriteLine();
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine("Invalid Option.....");
                                    Console.WriteLine();
                                    Console.ResetColor();
                                    break;
                            }
                            break;
                        case "3": /*FILES AND FOLDERS*/
                            Console.WriteLine("--------------------");
                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.Write(new string('\n', 2));
                            Console.WriteLine("o Choose Option");
                            Console.WriteLine();
                            option = @"
1. Upload a File
2. Downlad a File
3. Copy File(Between Libraries)
4. Delete File
5. Get All File Names and Folder Names
6. Create Folder
7. Exit
";
                            Console.WriteLine(option);
                            Console.ResetColor();
                            Console.Write("Option: ");
                            switch (Console.ReadLine())
                            {
                                case "1": /*UPLOAD A FILE*/
                                    Console.WriteLine();
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    Console.Write(new string('\n', 2));
                                    Console.Write("Please! mention the file path, which need to upload: ");
                                    Console.ResetColor();
                                    string uploadFilePath = "";
                                    uploadFilePath = Console.ReadLine();
                                    Console.WriteLine();

                                    await spApiCaller.UploadFile("https://graph.microsoft.com/v1.0/drives/" + config.LibraryId + "/root:/" + Path.GetFileName(uploadFilePath) + ":/content", result.AccessToken, uploadFilePath, Display);
                                    break;
                                case "2": /*DOWNLOAD A FILE*/
                                    Console.WriteLine();
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    Console.Write(new string('\n', 2));
                                    Console.Write("Please! mention the file name, which you want to download: ");
                                    Console.ResetColor();
                                    string downloadfileName = "";
                                    downloadfileName = Console.ReadLine();
                                    Console.WriteLine();

                                    await spApiCaller.DownloadFile("https://graph.microsoft.com/v1.0/drives/" + config.LibraryId + "/root:/" + downloadfileName + ":/content", result.AccessToken);
                                    break;
                                case "3": /*COPY FILE(BETWEEN LIBRARIES)*/
                                    string copyFileName = "", destinationLibraryName = "";
                                    Console.WriteLine();
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    Console.Write(new string('\n', 2));
                                    Console.Write("Please! mention the file name, which you want to copy: ");
                                    copyFileName = Console.ReadLine();
                                    Console.Write("Please! mention the library name, where you want to copy: ");
                                    destinationLibraryName = Console.ReadLine();
                                    Console.ResetColor();
                                    Console.WriteLine();

                                    await spApiCaller.CopyFile("https://graph.microsoft.com/v1.0/drives?$select=id,name", result.AccessToken, config.LibraryId, copyFileName, destinationLibraryName, Display);
                                    break;
                                case "4": /*DELETE FILE*/
                                    Console.WriteLine();
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    Console.Write(new string('\n', 2));
                                    Console.Write("Please! mention the file name, which you want to delete: ");
                                    Console.ResetColor();
                                    string deleteFileName = "";
                                    deleteFileName = Console.ReadLine();
                                    Console.WriteLine();

                                    await spApiCaller.DeleteFile("https://graph.microsoft.com/v1.0/drives/" + config.LibraryId + "/root:/" + deleteFileName, result.AccessToken, Display);
                                    break;
                                case "5": /*GET ALL FILE NAMES AND FOLDER NAMES*/
                                    Console.WriteLine();
                                    await spApiCaller.GetAllFilesandFolders("https://graph.microsoft.com/v1.0/drives/" + config.LibraryId + "/root/children?$select=name", result.AccessToken, Display);
                                    Console.WriteLine();
                                    break;
                                case "6": /*CREATE FOLDER*/
                                    Console.WriteLine();
                                    Console.ForegroundColor = ConsoleColor.Yellow;
                                    Console.Write(new string('\n', 2));
                                    Console.Write("Please! mention the folder name, which you want to create: ");
                                    Console.ResetColor();
                                    string FolderName = "";
                                    FolderName = Console.ReadLine();
                                    Console.WriteLine();

                                    var driveItem = new DriveItem
                                    {
                                        Name = FolderName,
                                        Folder = new Folder
                                        {
                                        },
                                        AdditionalData = new Dictionary<string, object>
    {
        { "@microsoft.graph.conflictBehavior", "rename" }
    },
                                    };
                                    JObject body = (JObject)JToken.FromObject(driveItem);
                                    await spApiCaller.CreateFolder("https://graph.microsoft.com/v1.0/drives/" + config.LibraryId + "/root/children", result.AccessToken, body, Display);

                                    break;
                                case "7": /*EXIT*/
                                    Environment.Exit(0);
                                    break;
                                default: /*INVALID OPTION*/
                                    Console.WriteLine();
                                    Console.ForegroundColor = ConsoleColor.Red;
                                    Console.WriteLine("Invalid Option.....");
                                    Console.WriteLine();
                                    Console.ResetColor();
                                    break;
                            }
                            break;
                        case "4": /*EXIT*/
                            Environment.Exit(0);
                            break;
                        default: /*INVALID OPTION*/
                            Console.WriteLine();
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("Invalid Option.....");
                            Console.WriteLine();
                            Console.ResetColor();
                            break;
                    }
                }
                #endregion
            }
        }

        /// <summary>
        /// Display the result of the Web API call
        /// </summary>
        /// <param name="result">Object to display</param>
        private static void Display(JObject result)
        {
            if (result != null)
            {
                foreach (JProperty child in result.Properties().Where(p => !p.Name.StartsWith("@")))
                {
                    Console.WriteLine($"{child.Name} = {child.Value}");
                }
            }
        }

#if VariationWithCertificateCredentials
        private static X509Certificate2 ReadCertificate(string certificateName)
        {
            if (string.IsNullOrWhiteSpace(certificateName))
            {
                throw new ArgumentException("certificateName should not be empty. Please set the CertificateName setting in the appsettings.json", "certificateName");
            }
            X509Certificate2 cert = null;

            using (X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser))
            {
                store.Open(OpenFlags.ReadOnly);
                X509Certificate2Collection certCollection = store.Certificates;

                // Find unexpired certificates.
                X509Certificate2Collection currentCerts = certCollection.Find(X509FindType.FindByTimeValid, DateTime.Now, false);

                // From the collection of unexpired certificates, find the ones with the correct name.
                X509Certificate2Collection signingCert = currentCerts.Find(X509FindType.FindBySubjectDistinguishedName, certificateName, false);

                // Return the first certificate in the collection, has the right name and is current.
                cert = signingCert.OfType<X509Certificate2>().OrderByDescending(c => c.NotBefore).FirstOrDefault();
            }
            return cert;
        }
#endif
    }
}
