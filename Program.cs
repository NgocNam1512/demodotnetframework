using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reconstructor
{
    class Program
    {
        public static void Main(string[] args)
        {
            //DocumentReconstructor.CreateDocument(args[0], args[1]);
            DocumentReconstructor.CreateDocument("data_file.json", "data_file.docx");

        }
    }
}
