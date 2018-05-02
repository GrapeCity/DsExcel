using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace GrapeCity.Documents.Excel.Examples
{
    public class Examples
    {
        private static FolderExample _rootExample;

        static Examples()
        {
            _rootExample = new RootExample(typeof(Examples).Namespace);

        }

        public static FolderExample RootExample
        {
            get
            {
                return _rootExample;
            }
        }
    }

    public class RootExample : FolderExample
    {
        public RootExample(string ns) : base(ns)
        {

        }

        protected override string NameResKey
        {
            get
            {
                return "RootExample.Name";
            }
        }

        protected override string DescripResKey
        {
            get
            {
                return "RootExample.Descrip";
            }
        }
    }

}
