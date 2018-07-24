using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelAddIn1
{
    class Constants
    {
        public const string PROTECTED_ERROR_MESSAGE = "Add-in has no permission to modify WorkBook's structure.";
        public const string VISIBLE_SHEET_LESS_THAN_TWO_MESSAGE = "visible sheet should be at least more than one.";
        public const string PermissionOperation = "Writable,ReadOnly,Invisible";
        public const string Writable = "Writable";
        public const string ReadOnly = "ReadOnly";
        public const string Invisible = "Invisible";
        public const string UserPasswordTable = "UserPasswordTable";
        public const string UserPermissionTable = "UserPermissionTable";
        public const string Mutable = "Mutable";
        public const string InMutable = "Inmutable";
        public const string key = "1234";
        public const string root = "root";
        public const string guest = "guest";
        public const string structure = "structure";
    }
}
