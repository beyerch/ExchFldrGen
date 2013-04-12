using System;
using System.IO;
using Microsoft.Exchange.WebServices.Data;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.Serialization.Json;
using System.Reflection;

namespace ExchangeUtilities
{
    class FolderGenerator
    {


        
        static void Main(string[] args)
        {

            //Create work variables
            string _evalString = "";


            //Display start
            Console.WriteLine("\n\nExchange Folder Generator v{0} ", Assembly.GetExecutingAssembly().GetName().Version);
            Console.WriteLine("----------------------------------------------------------------");

            //Create Command Line Dictionary
            Dictionary<string, string> _cmdLine = args.ParseCMDLineArgs("-");

            //Verify that we have all necessary command line parameters and that the ones supplied have valid input
            if (ValidateCommandLine(_cmdLine) == false)
            {
                Console.WriteLine();
                PrintSyntax();
                return;
            }


            //TO DO: Externalize these items; however, for now, use predefined values.
            //Dictionary<string, ExtendedPropertyDefinition> dctExtendedProperties = new Dictionary<string, ExtendedPropertyDefinition>();

            //If GUID values were supplied, load them 
            _cmdLine.TryGetValue("G", out _evalString);
            Dictionary<string, Guid> _dctGUIDs = new Dictionary<string, Guid>();
            if (_evalString != "" && _evalString != null)
            {
                _dctGUIDs = initGUIDs(_evalString);
            }

            //Create a list of mailboxes
            List<string> _dctMailboxes = GetMailboxes(_cmdLine);

            //Create the list of folders
            Dictionary<string, Folders> _dctFolders = GetFolders(_cmdLine["F"]);

            
            //It's go time, create some folders and apply retention tags
            exchCreateRetentionFolders(_cmdLine["VER"], _cmdLine["U"], _cmdLine["P"], _cmdLine["URL"], _dctMailboxes, _dctFolders, _dctGUIDs);


        }
        

        static void exchCreateRetentionFolders(string strVersion, string strIuser, string strIpwd, string strIURL, List<string> lstMailboxes, Dictionary<string, Folders> dctFolders, Dictionary<string, Guid> dctGUIDs)
        {
            try
            {
                //TO DO: Externalize these items; however, for now, use predefined values.
                // Create Policy Tag Information
                //Retention Policy Flag x3019
                ExtendedPropertyDefinition PolicyTag = new Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x3019, Microsoft.Exchange.WebServices.Data.MapiPropertyType.Binary);
                // PR_RETENTION_PERIOD 0x301A
                ExtendedPropertyDefinition RetentionFlags = new Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition(0x301D, Microsoft.Exchange.WebServices.Data.MapiPropertyType.Integer);
                

                // Connect to Exchange Web Services 
                ExchangeService service;
                switch (strVersion)
                {
                    case "2007SP1":
                        service = new ExchangeService(ExchangeVersion.Exchange2007_SP1);
                        break;
                    case "2010":
                        service = new ExchangeService(ExchangeVersion.Exchange2010);
                        break;
                    case "2010SP1":
                        service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
                        break;
                    case "2010SP2":
                        service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
                        break;
                    case "2013":
                        service = new ExchangeService(ExchangeVersion.Exchange2013);
                        break;
                    default:
                        service = new ExchangeService(ExchangeVersion.Exchange2010_SP1);
                        break;

                }

                try
                {
                    service.Credentials = new WebCredentials(strIuser, strIpwd);
                    service.AutodiscoverUrl(strIURL);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error connecting to Exchange Server : " + ex.Message);
                    return;
                }


                //Loop for people
                foreach (string _mailbox in lstMailboxes)
                {
                    Console.WriteLine("Processing Mailbox : {0}", _mailbox);
                    try
                    {
                        service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, _mailbox);

                        //Iterate through folders

                        //Start Folder / Policy creation process 
                        Console.WriteLine("\tCreating Folders...");

                        //Create working storage for Folder ID
                        ExtendedPropertyDefinition fldrRetentionPolicyTag = null;
                        System.Guid fldrRetentionPolicyGUID = Guid.NewGuid();
                        ExtendedPropertyDefinition fldrRetentionFlagTag = null;
                        int fldrRetentionFlagValue = 0;

                        foreach (KeyValuePair<string, Folders> _kvpFolder in dctFolders)
                        {
                            //Create working objects
                            Folders _folder = _kvpFolder.Value;
                            
                            //if FolderID and ParentID are not null, they are from a previous mailbox and are invalid, reset these values.
                            _folder.ParentID = null;
                            _folder.ID = null;

                            //Determine Retention Tag settings
                            //Dictionary<string, Folders> dctFolders, Dictionary<string, Guid> dctGUIDs
                            if (_folder.RetentionTagName != "" && _folder.RetentionTagName != null)
                                {
                                    //Search for the specified GUID and assign
                                    dctGUIDs.TryGetValue(_folder.RetentionTagName, out fldrRetentionPolicyGUID);
                                    //Produce warning if match wasn't found
                                    if (fldrRetentionPolicyGUID == Guid.Empty) Console.WriteLine("\t\tWARNING: Folder {0} has specified a retention flag of {1} which could not be located.  Ignoring retention setting.\n\t\tIf you supplied a GUID file (-g), ensure that retention GUIDs are correct.", _folder.FolderName, _folder.RetentionTagName);
                                }

                            if (fldrRetentionPolicyGUID != Guid.Empty)
                            {
                                fldrRetentionPolicyTag = PolicyTag;
                                fldrRetentionFlagTag = RetentionFlags;
                                fldrRetentionFlagValue = 145;
                            }
                            else
                            {
                                fldrRetentionPolicyTag = null;
                                fldrRetentionFlagTag = null;
                                fldrRetentionFlagValue = 0;
                            }

                            //Determine Parent Folder ID
                            switch (_folder.ParentFolderName)
                            {
                                case "WellKnownFolderName.ArchiveDeletedItems":
                                    _folder.ParentID = WellKnownFolderName.ArchiveDeletedItems;
                                    break;
                                case "WellKnownFolderName.ArchiveMsgFolderRoot":
                                    _folder.ParentID = WellKnownFolderName.ArchiveMsgFolderRoot;
                                    break;
                                case "WellKnownFolderName.ArchiveRecoverableItemsDeletions":
                                    _folder.ParentID = WellKnownFolderName.ArchiveRecoverableItemsDeletions;
                                    break;
                                case "WellKnownFolderName.ArchiveRecoverableItemsPurges":
                                    _folder.ParentID = WellKnownFolderName.ArchiveRecoverableItemsPurges;
                                    break;
                                case "WellKnownFolderName.ArchiveRecoverableItemsRoot":
                                    _folder.ParentID = WellKnownFolderName.ArchiveRecoverableItemsRoot;
                                    break;
                                case "WellKnownFolderName.ArchiveRecoverableItemsVersions":
                                    _folder.ParentID = WellKnownFolderName.ArchiveRecoverableItemsVersions;
                                    break;
                                case "WellKnownFolderName.ArchiveRoot":
                                    _folder.ParentID = WellKnownFolderName.ArchiveRoot;
                                    break;
                                case "WellKnownFolderName.DeletedItems":
                                    _folder.ParentID = WellKnownFolderName.DeletedItems;
                                    break;
                                case "WellKnownFolderName.Drafts":
                                    _folder.ParentID = WellKnownFolderName.Drafts;
                                    break;
                                case "WellKnownFolderName.Inbox":
                                    _folder.ParentID = WellKnownFolderName.Inbox;
                                    break;
                                case "WellKnownFolderName.Journal":
                                    _folder.ParentID = WellKnownFolderName.Journal;
                                    break;
                                case "WellKnownFolderName.JunkEmail":
                                    _folder.ParentID = WellKnownFolderName.JunkEmail;
                                    break;
                                case "WellKnownFolderName.LocalFailures":
                                    _folder.ParentID = WellKnownFolderName.LocalFailures;
                                    break;
                                case "WellKnownFolderName.MsgFolderRoot":
                                    _folder.ParentID = WellKnownFolderName.MsgFolderRoot;
                                    break;
                                case "WellKnownFolderName.Notes":
                                    _folder.ParentID = WellKnownFolderName.Notes;
                                    break;
                                case "WellKnownFolderName.Outbox":
                                    _folder.ParentID = WellKnownFolderName.Outbox;
                                    break;
                                case "WellKnownFolderName.PublicFoldersRoot":
                                    _folder.ParentID = WellKnownFolderName.PublicFoldersRoot;
                                    break;
                                case "WellKnownFolderName.QuickContacts":
                                    _folder.ParentID = WellKnownFolderName.QuickContacts;
                                    break;
                                case "WellKnownFolderName.RecipientCache":
                                    _folder.ParentID = WellKnownFolderName.RecipientCache;
                                    break;
                                case "WellKnownFolderName.RecoverableItemsDeletions":
                                    _folder.ParentID = WellKnownFolderName.RecoverableItemsDeletions;
                                    break;
                                case "WellKnownFolderName.RecoverableItemsPurges":
                                    _folder.ParentID = WellKnownFolderName.RecoverableItemsPurges;
                                    break;
                                case "WellKnownFolderName.RecoverableItemsRoot":
                                    _folder.ParentID = WellKnownFolderName.RecoverableItemsRoot;
                                    break;
                                case "WellKnownFolderName.RecoverableItemsVersions":
                                    _folder.ParentID = WellKnownFolderName.RecoverableItemsVersions;
                                    break;
                                case "WellKnownFolderName.Root":
                                    _folder.ParentID = WellKnownFolderName.Root;
                                    break;
                                case "WellKnownFolderName.SearchFolders":
                                    _folder.ParentID = WellKnownFolderName.SearchFolders;
                                    break;
                                case "WellKnownFolderName.SentItems":
                                    _folder.ParentID = WellKnownFolderName.SentItems;
                                    break;
                                case "WellKnownFolderName.ServerFailures":
                                    _folder.ParentID = WellKnownFolderName.ServerFailures;
                                    break;
                                case "WellKnownFolderName.SyncIssues":
                                    _folder.ParentID = WellKnownFolderName.SyncIssues;
                                    break;
                                case "WellKnownFolderName.Tasks":
                                    _folder.ParentID = WellKnownFolderName.Tasks;
                                    break;
                                case "WellKnownFolderName.ToDoSearch":
                                    _folder.ParentID = WellKnownFolderName.ToDoSearch;
                                    break;
                                case "WellKnownFolderName.VoiceMail":
                                    _folder.ParentID = WellKnownFolderName.VoiceMail;
                                    break;
                                default:
                                    //Not a system folder, so we need to go through the folders dictionary and find the ParentID
                                    Folders _parentfolder;

                                    dctFolders.TryGetValue(_folder.ParentFolderName, out _parentfolder);
                                    if (_parentfolder != null) _folder.ParentID = _parentfolder.ID;
                                    break;
                            }

                            
                            if (_folder.ParentID != null)
                            {
                                //Create the folder entry
                                /*
                                Microsoft.Exchange.WebServices.Data.FolderId fldrParentID;
                                ExtendedPropertyDefinition fldrRetentionPolicyTag;
                                System.Guid fldrRetentionPolicyGUID;
                                ExtendedPropertyDefinition fldrRetentionFlagTag;
                                int fldrRetentionFlagValue;
                                */

                                _folder.ID = exchCreateFolder(service, _folder.FolderName, _folder.ParentID, fldrRetentionPolicyTag, fldrRetentionPolicyGUID, fldrRetentionFlagTag, fldrRetentionFlagValue);

                            }
                            else
                            {
                                //Produce warning if match wasn't found
                                Console.WriteLine("\t\tWARNING: Skipping folder {0}, unable to determine parent folder ID!", _folder.FolderName);
                            }

                        }

                        // Write confirmation message to console window.
                        Console.WriteLine("Folders created successfully for " + _mailbox);

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error impersonating mailbox : " + _mailbox + ", Error : " + ex.Message);
                        return;
                    }


                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error executing program : " + ex.Message);
            }

       
        }

        static Microsoft.Exchange.WebServices.Data.FolderId exchCreateFolder(ExchangeService service, string fldrDisplayName, Microsoft.Exchange.WebServices.Data.FolderId fldrParentID, ExtendedPropertyDefinition fldrRetentionPolicyTag, System.Guid fldrRetentionPolicyGUID, ExtendedPropertyDefinition fldrRetentionFlagTag, int fldrRetentionFlagValue)
        {
            // Verify that the paramters passed in are valid
            if (fldrDisplayName.Length == 0)
            {
                Console.WriteLine("\tError in exchCreateFolder : Invalid Folder Name");
                return null;
            }

            //Generate process output          
            Console.Write("\tCreating Folder : " + fldrDisplayName + "...");            
            
            if (service == null)
            {
                Console.WriteLine("error:Exchange Service reference Invalid");
                return null;
            }
            if (fldrParentID == null)
            {
                Console.WriteLine("error:Invalid Parent Folder");
                return null;
            }
            //if (fldrRetentionPolicyTag == null)
            //{
            //    Console.WriteLine("\t\tError in exchCreateFolder : Invalid Retention Policy Tag");
            //    return null;
            //}
            //if (fldrRetentionPolicyGUID == null)
            //{
            //    Console.WriteLine("\tError in exchCreateFolder : Invalid Retention Policy GUID");
            //    return null;
            //}

            //if (fldrRetentionFlagTag == null)
            //{
            //    Console.WriteLine("\t\tError in exchCreateFolder : Invalid Retention Policy Flag Tag");
            //    return null;
            //}


            //if (fldrRetentionFlagValue < 1)
            //{
            //    Console.WriteLine("\t\tError in exchCreateFolder : Invalid Retention Policy Flag Value");
            //    return null;
            //}
        
            //We have all of the inforamtion we need to make folders, create folder and assign retention policy.

            // Create Folder instance 
            Folder newFolder = new Folder(service);

            try
            {
                // Assign attributes to folder and save it to exchange server
                newFolder.DisplayName = fldrDisplayName;
                if (fldrRetentionPolicyTag != null)
                {
                    Console.Write("Applying Retention Policy Tag..");
                    newFolder.SetExtendedProperty(fldrRetentionPolicyTag, fldrRetentionPolicyGUID.ToByteArray());
                    newFolder.SetExtendedProperty(fldrRetentionFlagTag, fldrRetentionFlagValue);
                    
                }
                newFolder.Save(fldrParentID);

            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : " + ex.Message);
                return null;

            }

            Console.WriteLine("success.");
            //Return the folder id reference
            return newFolder.Id;
        }

        static void PrintSyntax()
        {
            // Display syntax information 
            Console.WriteLine("Communicates with Exchange Server and creates predetermined folder structures for 1 to n users and optionally applies extended properties\n");
            Console.WriteLine("USAGE:");
            Console.WriteLine("\tExchFldrGen.exe [-ver <2007SP1|2010|2010SP1|2010SP2|2013>] -u <mailbox user> -p <password> \n\t\t[-url <AutoDiscover URL>] {-e <email address> | -ef <email address file>} -f <folder list file> [-g <guid file>]  \n");
            Console.WriteLine("Options:");
            Console.WriteLine("\t-ver\t\tSpecifies the Exchange server version.  If not specified, default is 2010SP1.");
            Console.WriteLine("\t-u\t\tImpersonation account that has mailbox rights to all specified mailboxes.\n\t\t\t(If updating one mailbox, can use mailbox credentials)");
            Console.WriteLine("\t-p\t\tAccount password");
            Console.WriteLine("\t-url\t\tURL/Email address used to discover Exchange server.  If not supplied will attempt to use Impersonator address");
            Console.WriteLine("\t-e | -ef\tEmail address options: (either -e or -ef must be specified)");
            Console.WriteLine("\t\t-e\t\tExecute Retention Folder creation by specified mailbox");
            Console.WriteLine("\t\t-ef\t\tExecute Retention Folder creation by specified mailbox list file");
            Console.WriteLine("\t-f\t\tFile containing list of folders, folder structure, and Extended Property tags.");
            Console.WriteLine("\t-g\t\tSpecify file that contains GUIDs for Extended Properties [OPTIONAL if not using Extended Properties]\t");
            Console.WriteLine("where");
            Console.WriteLine("\t<mailbox user>\tUser who has Impersonation/Mailbox rights for all mailboxes being updated i.e. admin@test.com");
            Console.WriteLine("\t<password>\tAccount password");
            Console.WriteLine("\t<Auto URL>\tAccount/URL which can be used to determine Exchange server addressAccount password");
            Console.WriteLine("\t<GUIDfile>\tFile containing GUIDs for Extended Properties");
            Console.WriteLine("\t<email address>\tExchange Mailbox SMTP account");
            Console.WriteLine("\t<email address file>\tText file containing one mailbox per row that will be processed.");
            Console.WriteLine("\t<folder list file>\tStructured XML file containing folder information, see sample file.");
            Console.WriteLine("\t<guid file>\tJSON file listing Extended Property GUIDs if Structured XML file containing folder information, see sample file.\n");
            Console.WriteLine("Examples:\n\n> ExchFldrGen.exe -ver2007SP1 -u admin@foo.com -p 1234 -e john.doe@foo.com -f folders.xml\n\t- Adds the folders as specified in folders.xml to john doe's mailbox on the foo.com Exchange 2007 SP1 server");
            Console.WriteLine("Examples:\n\n> ExchFldrGen.exe -u admin@foo.com -p 1234 -ef mailboxes.csv -f folders.xml -g GUIDs.json\n\t- For each mailbox (on the foo.com Exchange 2010 SP1 server) specified in the file mailboxes.csv, adds the folders as \n\t  specified in folders.xml and applies the extended properties as defined in folders.XML and GUIDs.json. ");
            Console.WriteLine("\nNOTES:");
            Console.WriteLine("\nIn order for this program to work, the Exchange Server must have the EWS Managed API installed. \nSee : http://msdn.microsoft.com/en-us/library/exchange/dd633710%28v=exchg.80%29.aspx");

        }

        static bool ValidateCommandLine(Dictionary<string, string> _cmdLine)
        {

            bool boolValidParameters = true;
            bool boolValidParamE = false;
            bool boolValidParamEF = false;

            string _evalString = "";

            //Check Version and set default if notpresent
            try
            {
                switch (_cmdLine["VER"])
                {
                    case "2007SP1":
                    case "2010":
                    case "2010SP1":
                    case "2010SP2":
                    case "2013":
                        break;
                    default:
                        Console.WriteLine("WARNING: Invalid parameter value, {0}, for -ver.  Using Default value 2010SP1", _cmdLine["VER"]);
                        _cmdLine["VER"] = "2010SP1";
                        break;
                }

            }
            catch (KeyNotFoundException)
            {
                Console.WriteLine("INFORMATIONAL: -ver parameter not supplied, defaulting to 2010SP1");
                _cmdLine.Add("VER", "2010SP1");
            }

            //Check to see if user / impersonal account passed
            _cmdLine.TryGetValue("U", out _evalString);
            if (_evalString == "" || _evalString == null)
            {
                Console.WriteLine("ERROR: -u parameter is missing.");
                boolValidParameters = false;
            }

            //Check to see if user / impersonal account passed
            _cmdLine.TryGetValue("P", out _evalString);
            if (_evalString == "" || _evalString == null)
            {
                Console.WriteLine("ERROR: -p parameter is missing.");
                boolValidParameters = false;
            }

            //Check to see if autodiscover passed
            if (_cmdLine.TryGetValue("URL", out _evalString))
            {
                if (_evalString == "" || _evalString == null)
                {
                    try
                    {
                        Console.WriteLine("INFORMATIONAL: -url parameter is missing, using -u value ({0}) instead.", _cmdLine["U"]);
                        _cmdLine["URL"] = _cmdLine["U"];
                    }
                    catch (KeyNotFoundException)
                    {
                        Console.WriteLine("ERROR: -url parameter is missing, and cannot default to -u value");
                    }

                }
            }
            else
            {

                if (_evalString == "" || _evalString == null)
                {
                    try
                    {
                        Console.WriteLine("INFORMATIONAL: -url parameter is missing, using -u value ({0}) instead.", _cmdLine["U"]);
                        _cmdLine.Add("URL", _cmdLine["U"]);
                    }
                    catch (KeyNotFoundException)
                    {
                        Console.WriteLine("ERROR: -url parameter is missing, and cannot default to -u value");
                    }

                }
            }

            //Check to see if -e sent
            _cmdLine.TryGetValue("E", out _evalString);
            if (_evalString != "" && _evalString != null)
            {
                boolValidParamE = true;
            }

            if (boolValidParamE == false)
            {
                //Check to see if -ef sent
                _cmdLine.TryGetValue("EF", out _evalString);
                if (_evalString != "" && _evalString != null)
                {
                    //Attempt to validate file passed
                    if (System.IO.File.Exists(_evalString))
                    {
                        boolValidParamEF = true;
                    }
                    else
                    {
                        Console.WriteLine("ERROR: -ef file [{0}] does not exist!", _evalString);
                    }

                }
            }

            //check to make sure either -e or -ef was valid
            if (!boolValidParamE &&  !boolValidParamEF)
            {
                boolValidParameters = false;
                Console.WriteLine("ERROR: Either the -e or -ef parameter must be provided.");
            }


            //Check to see if -f sent
            _cmdLine.TryGetValue("F", out _evalString);
            if (_evalString != "" && _evalString != null)
            {
                //Attempt to validate file passed
                if (!System.IO.File.Exists(_evalString))
                {
                    boolValidParameters = false;
                    Console.WriteLine("ERROR: -f file [{0}] does not exist!", _evalString);
                }
            }
            else
            {
                boolValidParameters = false;
                Console.WriteLine("ERROR: -f parameter not specified.", _evalString);
            }


            //Check to see if -g sent
            _cmdLine.TryGetValue("G", out _evalString);
            if (_evalString != "" && _evalString != null)
            {
                //Attempt to validate file passed
                if (!System.IO.File.Exists(_evalString))
                {
                    boolValidParameters = false;
                    Console.WriteLine("ERROR: -g file [{0}] does not exist!", _evalString);
                }
            }


            return boolValidParameters;
        }


        static Dictionary<string, Guid> initGUIDs(string strFileName)
        {
            Dictionary<string, Guid> _dctGUIDs = new Dictionary<string, Guid>();

            try
            {

                if (!System.IO.File.Exists(strFileName))
                {
                    using (System.IO.FileStream fs = System.IO.File.Create(strFileName))
                    {

                        //Serialize defaults  
                        //TO DO: Remove this from final production, only used internally to build the first JSON file so I didn't have to construct it manually....
                        _dctGUIDs.Clear();
                        _dctGUIDs.Add("Never", new Guid("{9971ea35-e10b-4fa1-a25c-bfbaade6bced}"));
                        _dctGUIDs.Add("Default", new Guid("{ca5fddc1-a645-4f5a-a8bf-4ef514f75f2a}"));
                        _dctGUIDs.Add("DeletedItems", new Guid("{3a7e439c-6bb7-464f-9d4c-5f842f4484fc}"));
                        _dctGUIDs.Add("JunkEmail", new Guid("{dbde7253-78d7-4cd5-a432-6b607e36059f}"));
                        _dctGUIDs.Add("1Year", new Guid("{9c9ca00b-ce70-4467-a6a4-30a8db3c89e2}"));
                        _dctGUIDs.Add("2Years", new Guid("{35180183-ce65-4461-9312-4e3da1a87bb8}"));
                        _dctGUIDs.Add("3Years", new Guid("{86c0cd9e-1cc5-449a-870f-92e577be72d1}"));
                        _dctGUIDs.Add("4Years", new Guid("{37a9d4d2-867a-4f96-b95d-a9c9267fbdab}"));
                        _dctGUIDs.Add("5Years", new Guid("{77602efc-c80f-4aa6-8e15-efff312ac5fd}"));
                        _dctGUIDs.Add("6Years", new Guid("{329b16a1-a532-4086-8f37-a3343e5d0679}"));
                        _dctGUIDs.Add("7Years", new Guid("{4c5edebd-d6fa-48de-921e-bea4d94bad67}"));
                        _dctGUIDs.Add("8Years", new Guid("{aa1d6fdc-a215-4bf2-893c-b2980925c9c9}"));
                        _dctGUIDs.Add("9Years", new Guid("{d4da2d2c-51c0-4880-9c2f-bf93e01c6bd2}"));
                        _dctGUIDs.Add("10Years", new Guid("{44cd0f53-c58f-4314-b0ce-ae9d24513abb}"));
                        _dctGUIDs.Add("11Years", new Guid("{4b09947e-3ec3-4276-b277-7456171fa7b2}"));
                        _dctGUIDs.Add("15Years", new Guid("{d3aa0551-2149-44f8-a549-8cdf1bf88919}"));
                        _dctGUIDs.Add("20Years", new Guid("{4c51458f-af3e-4a01-919e-013fd4288de2}"));
                        _dctGUIDs.Add("16Years", new Guid("{70dcf66e-7a7a-4f74-9e40-c10513d1f3fd}"));
                        _dctGUIDs.Add("30Years", new Guid("{833f95a3-eb3f-4934-9371-78d3404f9ad6}"));
                        _dctGUIDs.Add("15Months", new Guid("{81000566-534e-4eb6-81f0-7cdd21ff87bb}"));
                        _dctGUIDs.Add("18Months", new Guid("{7af3d385-8d03-4dd3-81e6-b6f115a7def7}"));
                        _dctGUIDs.Add("6Months", new Guid("{f20e0875-7ab1-42d8-a0bb-efcf5d2ed726}"));
                        _dctGUIDs.Add("DepartmentMailboxFolders", new Guid("{ec8447eb-7bdd-4a46-8645-b272bbd2556e}"));
                        _dctGUIDs.Add("SentItems", new Guid("{bb0632a3-60a0-43ab-8596-61d020a4a410}"));



                        //Create the file for future reference
                        using (MemoryStream stream = new MemoryStream())
                        {
                            DataContractJsonSerializer s1 = new DataContractJsonSerializer(typeof(Dictionary<string, Guid>));

                            s1.WriteObject(stream, _dctGUIDs);
                            stream.WriteTo(fs);
                        }

                        //Close the File Stream
                        fs.Close();

                    }
                }
                else
                {
                    using (System.IO.FileStream fs = System.IO.File.OpenRead(strFileName))
                    {

                        //Create the file for future reference
                        using (MemoryStream stream = new MemoryStream())
                        {
                            stream.SetLength(fs.Length);
                            fs.Read(stream.GetBuffer(), 0, (int)fs.Length);
                            stream.Flush();

                            DataContractJsonSerializer s1 = new DataContractJsonSerializer(typeof(Dictionary<string, Guid>));
                            _dctGUIDs = (Dictionary<string, Guid>)s1.ReadObject(stream);

                        }

                        foreach (KeyValuePair<string, Guid> _guids in _dctGUIDs)
                        {
                            Console.WriteLine("Key : \"{0}\", Type= \"{1}\", Value : \"{2}\", Type= \"{3}\".", _guids.Key.ToString(), _guids.Key.GetType().ToString(), _guids.Value.ToString(), _guids.Value.GetType().ToString());

                        }

                        //Close the File Stream
                        fs.Close();

                    }

                    Console.WriteLine("File \"{0}\" already exists.", strFileName);

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error encountered retrieving GUIDs : Msg = {0}", ex.Message);
            }

            return _dctGUIDs;
        }

        static List<string> GetMailboxes(Dictionary<string, string> _cmdLine)
        {
            //Declare working variables
            string _evalString = "";
            List<string> _dctMailboxes = new List<string>();
            bool inDebug = true;



            _cmdLine.TryGetValue("E", out _evalString);
            if (_evalString != "" && _evalString != null)
            {
                _dctMailboxes.Add(_evalString);
            }
            else
            {
                _cmdLine.TryGetValue("EF", out _evalString);
                try
                {
                    using (System.IO.StreamReader fs = new System.IO.StreamReader(_evalString))
                    {
                        string _linedata = "";
                        while ((_linedata = fs.ReadLine()) != null)
                        {
                            _dctMailboxes.Add(_linedata.Trim().ToString());
                        }

                        //Done close the file
                        fs.Close();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error encountered retrieving Mailboxes from file {0}.  Error Mesg = {1}", _evalString, ex.Message);
                }
            }

            //For Debug only
            if (inDebug)
            {
                Console.WriteLine("DEBUG: Processing the following mailboxes");
                foreach (string _mailbox in _dctMailboxes)
                {
                    Console.WriteLine("\t Adding mailbox : {0}.", _mailbox.ToString());
                }
            }

            //Return mailbox list
            return _dctMailboxes;

        }


        static Dictionary<string, Folders> GetFolders(string _mailboxFile)
        {
            //Declare working variables
            Dictionary<string, Folders> _dctFolders = new Dictionary<string, Folders>();
            bool inDebug = true;

            if (!System.IO.File.Exists(_mailboxFile))
            {
                //Create a Sample File
                using (System.IO.FileStream fs = System.IO.File.Create("folders.sam"))
                {

                    //Serialize defaults  
                    //TO DO: Remove this from final production, only used internally to build the first JSON file so I didn't have to construct it manually....

                    //Create the list of folders
                    _dctFolders.Add("Test 1", new Folders { FolderName = "Test 1", ParentFolderName = "WellKnownFolderName.MsgFolderRoot", RetentionTagName = "1Year", ID = null, ParentID = null });
                    _dctFolders.Add("Test 2", new Folders { FolderName = "Test 2", ParentFolderName = "Test 1", RetentionTagName = "1Year", ID = null, ParentID = null });
                    _dctFolders.Add("Test 3", new Folders { FolderName = "Test 3", ParentFolderName = "Test 2", RetentionTagName = "1Year", ID = null, ParentID = null });
                    _dctFolders.Add("Test 4", new Folders { FolderName = "Test 4", ParentFolderName = "Test 3", RetentionTagName = "1Year", ID = null, ParentID = null });

                    try
                    {
                        //Create the file for future reference
                        using (MemoryStream stream = new MemoryStream())
                        {
                            DataContractJsonSerializer s1 = new DataContractJsonSerializer(typeof(Dictionary<string, Folders>));

                            s1.WriteObject(stream, _dctFolders);
                            stream.WriteTo(fs);
                        }

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("ERROR: creating folders.sam, {0}", ex.Message);
                    }

                    //Close the File Stream
                    fs.Close();

                }
            }
            else
            {

                using (System.IO.FileStream fs = System.IO.File.OpenRead(_mailboxFile))
                {

                    //Create the file for future reference
                    using (MemoryStream stream = new MemoryStream())
                    {
                        stream.SetLength(fs.Length);
                        fs.Read(stream.GetBuffer(), 0, (int)fs.Length);
                        stream.Flush();

                        DataContractJsonSerializer s1 = new DataContractJsonSerializer(typeof(Dictionary<string, Folders>));
                        _dctFolders = (Dictionary<string, Folders>)s1.ReadObject(stream);

                    }


                    //Close the File Stream
                    fs.Close();

                }

            }

            //For Debug only
            if (inDebug)
            {
                Console.WriteLine("DEBUG: Processing the following folders");
                foreach (KeyValuePair<string, Folders> _folders in _dctFolders)
                {
                    Console.WriteLine("\tFolder Name = {0}, Parent Name = {1}, Retention Policy Tag = {2}", _folders.Value.FolderName, _folders.Value.ParentFolderName, _folders.Value.RetentionTagName);

                }
            }

            //Return mailbox list
            return _dctFolders;

        }


    }
}
