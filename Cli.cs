using OutlookAttachmentsDownloader.Exceptions;
using Sharprompt;
using System;
using System.Linq;
using System.Threading.Tasks;

namespace OutlookAttachmentsDownloader
{
    /// <summary>
    /// Implementation of cli tool for user interaction.
    /// </summary>
    public sealed class Cli
    {
        private static readonly Cli instance = new Cli();
        private readonly OApp oapp;
        private Cli() {
            oapp = OApp.Instance;
        }

        public static Cli Instance
        {
            get {  return instance; }
        }

        /// <summary>
        /// Begin the prompt sequence with the user.
        /// </summary>
        /// <returns></returns>
        public async Task  InitializeSequence()
        {
            AskEmailAccount();
            await AskFolders();
            AskDestination();
            AskConfirmation();
            await StartDownloads();

        }

        private void AskEmailAccount()
        {
            var email = Prompt.Select("Select the account", oapp.AccountsAvailable);
            oapp.SelectedAccount = email;
        }

        private async Task AskFolders()
        {
            if(oapp.SelectedAccount == null)
            {
                Console.WriteLine("No account selected ! ");
                AskEmailAccount();
            }
            string[] folderList = await oapp.FetchFolderList();
            var folders = Prompt.MultiSelect("Select all folders required",folderList, pageSize: 15);
            oapp.SelectedFolders = folders.ToArray();
        }

        private void AskDestination()
        {
            var destination = Prompt.Input<string>("Path","",new[] { Validators.Required() });
            oapp.Destination = destination;
        }

        private void AskConfirmation()
        {
            var answer = Prompt.Confirm("Please confirm all your previous choices", defaultValue: true);
            if(!answer)
            {
                Console.WriteLine("All right ! the program will exit");
                throw new AbortedOperationException("Program execution aborted by user.");
            }
        }

        private async Task StartDownloads()
        {
            Console.WriteLine("Downloading attachments...");
            await oapp.SaveAttachments();
            Console.WriteLine("Download complete");
        }
    }
}
