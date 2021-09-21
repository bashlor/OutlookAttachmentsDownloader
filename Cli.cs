using Sharprompt;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAttachmentsDownloader
{
    public sealed class Cli
    {
        private static readonly Cli instance = new Cli();
        private OApp oapp;
        private Cli() {
            oapp = OApp.Instance;
        }

        public static Cli Instance
        {
            get {  return instance; }
        }

        public async Task  InitializeSequence()
        {
            AskEmailAccount();
            AskFolders();
            AskDestination();
            AskConfirmation();
            await startDownloads();

        }

        private void AskEmailAccount()
        {
            var email = Prompt.Select("Select the account", oapp.AccountsAvailable);
            oapp.SelectedAccount = email;
            Console.WriteLine("You selected : " + email);
        }

        private void AskFolders()
        {
            if(oapp.SelectedAccount == null)
            {
                Console.WriteLine("No account selected ! ");
                AskEmailAccount();
            }
            var folders = Prompt.MultiSelect("Select all folders required", oapp.FoldersList, pageSize: 15);
            Console.WriteLine($"You selected {string.Join(", ", folders)}");
        }

        private void AskDestination()
        {
            var destination = Prompt.Input<string>("Which path ?");
            oapp.Destination =destination;
        }

        private void AskConfirmation()
        {
            var answer = Prompt.Confirm("Please confirm all your previous choices", defaultValue: true);
            if(answer == false)
            {
                Console.WriteLine("All right ! the program will exit");
                throw new SystemException("Program execution aborted by user.");
            }
        }

        private async Task startDownloads()
        {
            Console.WriteLine("Downloading attachments...");
            await oapp.SaveAttachments();
            Console.WriteLine("Download complete");
        }
    }
}
