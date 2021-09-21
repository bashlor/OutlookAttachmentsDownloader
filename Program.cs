using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAttachmentsDownloader
{
    class Program
    {
        static int Main()
        {
            Cli cli = Cli.Instance;
            Task task =  cli.InitializeSequence();
            task.Start();
            task.Wait();
            return 0;
        }
    }
}
