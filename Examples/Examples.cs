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

        public static ExampleBase[] GetAllExamples()
        {
            List<ExampleBase> examples = new List<ExampleBase>();
            examples.Add(RootExample);
            foreach (var child in RootExample.Children)
            {
                GetExamples(child, examples);
            }

            return examples.ToArray();
        }
       
        private static void GetExamples(ExampleBase example, List<ExampleBase> examples)
        {
            examples.Add(example);
            if (example is FolderExample)
            {
                FolderExample folderExample = example as FolderExample;
                foreach (var child in folderExample.Children)
                {
                    GetExamples(child, examples);
                }
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
