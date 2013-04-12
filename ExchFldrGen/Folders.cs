using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;


namespace ExchangeUtilities
{
    [DataContract]
    class Folders
    {
        [DataMember]
        public string FolderName {get; set;}
        [DataMember]
        public string ParentFolderName {get; set;}
        [DataMember]
        public string RetentionTagName {get; set;}
        [DataMember]
        public Microsoft.Exchange.WebServices.Data.FolderId ID {get; set;}
        [DataMember]
        public Microsoft.Exchange.WebServices.Data.FolderId ParentID {get; set;}

    }
}
