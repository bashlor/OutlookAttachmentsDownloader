using System;
using System.Collections.Generic;
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
            }catch(System.Runtime.InteropServices.COMException ex)
            {
                Console.WriteLine(ex.Message);
                return -1;
            }
            catch (SystemException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
                return 1;
            }
            return 0;
        }
    }
}
