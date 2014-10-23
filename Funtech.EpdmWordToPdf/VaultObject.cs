using EdmLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Funtech.EpdmWordToPdf
{
    public class VaultObject
    {
        public EdmObjectType ObjectType { get; set; }
        public int Id { get; set; }

        /// <summary>
        /// Only valid if object is a file (will be null for folders)
        /// </summary>
        public int? ParentFolderId { get; set; }
    }
}
