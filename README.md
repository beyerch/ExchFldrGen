Exchange Folder Generator v1.0.0.0
----------------------------------------------------------------

Communicates with Exchange Server and creates predetermined folder structures for 1 to n users and optionally applies 
extended properties

USAGE:
        ExchFldrGen.exe [-ver <2007SP1|2010|2010SP1|2010SP2|2013>] -u <mailbox user> -p <password>
                [-url <AutoDiscover URL>] {-e <email address> | -ef <email address file>} -f <folder list file> 
                [-g <guid file>]

Options:
        -ver            Specifies the Exchange server version.  If not specified, default is 2010SP1.
        -u              Impersonation account that has mailbox rights to all specified mailboxes.
                        (If updating one mailbox, can use mailbox credentials)
        -p              Account password
        -url            URL/Email address used to discover Exchange server.  If not supplied will attempt to use Impersonator address
        -e | -ef        Email address options: (either -e or -ef must be specified)
                -e              Execute Retention Folder creation by specified mailbox
                -ef             Execute Retention Folder creation by specified mailbox list file
        -f              File containing list of folders, folder structure, and Extended Property tags.
        -g              Specify file that contains GUIDs for Extended Properties [OPTIONAL if not using Extended Properties]
where
        <mailbox user>  User who has Impersonation/Mailbox rights for all mailboxes being updated i.e. admin@test.com
        <password>      Account password
        <Auto URL>      Account/URL which can be used to determine Exchange server addressAccount password
        <GUIDfile>      File containing GUIDs for Extended Properties
        <email address> Exchange Mailbox SMTP account
        <email address file>    Text file containing one mailbox per row that will be processed.
        <folder list file>      Structured XML file containing folder information, see sample file.
        <guid file>     JSON file listing Extended Property GUIDs if Structured XML file containing folder information, see sample file
.

Examples:

> ExchFldrGen.exe -ver2007SP1 -u admin@foo.com -p 1234 -e john.doe@foo.com -f folders.xml
        - Adds the folders as specified in folders.xml to john doe's mailbox on the foo.com Exchange 2007 SP1 server
Examples:

> ExchFldrGen.exe -u admin@foo.com -p 1234 -ef mailboxes.csv -f folders.xml -g GUIDs.json
        - For each mailbox (on the foo.com Exchange 2010 SP1 server) specified in the file mailboxes.csv, adds the folders as
          specified in folders.xml and applies the extended properties as defined in folders.XML and GUIDs.json.

NOTES:

In order for this program to work, the Exchange Server must have the EWS Managed API installed.
See : http://msdn.microsoft.com/en-us/library/exchange/dd633710%28v=exchg.80%29.aspx

C:\Users\cbeyer\Documents\Visual Studio 2010\Projects\ExchFldrGen\ExchFldrGen\ExchFldrGen\bin\Release>
