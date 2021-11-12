using System;
using System.Reflection;
using System.Threading.Tasks;

[assembly: AssemblyVersion("1.0.0.0")]
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
            OApp.Instance.CloseInstance();
            return 0;
        }
    }
}
