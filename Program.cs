using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAttachmentsDownloader
{
    class Program
    {
        static async Task<int> Main()
        {
            try
            {
                Cli cli = Cli.Instance;
                await cli.InitializeSequence();
                Console.ReadKey();
            }catch(System.Runtime.InteropServices.COMException ex)
            {
                Console.WriteLine(ex.Message);
                
            }
            catch (SystemException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
                return 1;
            }
            OApp.Instance.closeInstance();
            return 0;
        }
    }
}
