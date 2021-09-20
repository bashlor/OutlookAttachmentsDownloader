using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;


using System.Threading.Tasks;

namespace OutlookAttachmentsDownloader
{
    public sealed class OApp
    {
        public static OApp Instance { get { return instance; } }
        private static readonly OApp instance = new OApp();

        readonly NameSpace outlookNamespace;
        readonly Application outlookApplication;

        readonly Dictionary<string, Folder> accounts;


        string selectedAccount;
        string selectedFolder;
        string destination;

        static OApp() { }

        public string SelectedAccount
        {
            get {
                return selectedAccount;
            }
            set {  
                if(accounts.ContainsKey(value))
                {
                    selectedAccount = value;
                }
                else
                    throw new System.Exception("Account not found ! Pls check again !");
            }
        }

        public string SelectedFolder
        {
            get
            {
                return selectedFolder;
            }
            set
            {
                selectedFolder = value;
            }
        }

        public string Destination
        {
            get
            {
                return destination;
            }
            set
            {
                destination = value;
            }
        }

        public async void SaveAttachments()
        {
            Folder folder = await SearchFolder(selectedFolder, accounts[selectedAccount]);
            if(folder == null)
            {
                throw new System.Exception("Folder " + selectedFolder + " not found.");
            }
            else
            {
                Console.WriteLine("Folder " + selectedFolder +  " found ! ");
                await SaveAttachmentForEveryMailItem(folder);
                EnumerateFolders(folder, (subFolder) =>
                {
                    Console.WriteLine("Exploring " + subFolder.FolderPath + "...");
                    SaveAttachmentForEveryMailItem(folder);
                });
            }

        }

        private void EnumerateFolders(Folder folder, System.Action<Folder> callback)
        {
            Folders childFolders = folder.Folders;
            if (childFolders.Count > 0)
            {
                foreach (Folder childFolder in childFolders)
                {
                    callback(childFolder);
                    EnumerateFolders(childFolder,callback);
                }
            }

        }



        private void fetchAccountsList()
        {
            foreach (MAPIFolder folder in outlookApplication.Session.Folders)
            {
                if (!accounts.ContainsKey(folder.Name))
                {
                    accounts.Add(folder.Name, folder as Folder);
                }
            }
        }

        private OApp() {
            try
            {
                outlookApplication = new Application();
                outlookNamespace = outlookApplication.GetNamespace("MAPI");
                accounts = new Dictionary<string, Folder>();
                _initialization();

            }catch (System.Exception ex) {
                Console.WriteLine("OApp Outlook intiialization error : " + ex.Message);
                closeInstance();
            }
        }

        private void closeInstance()
        {
            if(outlookApplication != null)
            {
               //Release here COM objects
               foreach(KeyValuePair<string,Folder> account in accounts)
                {
                    ReleaseComObject(account.Value);
                }
                ReleaseComObject(outlookNamespace);
                ReleaseComObject(outlookApplication);
            }
        }

        private static void ReleaseComObject(object obj)
        {
            if (obj != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            }
        }

        private  void _initialization()
        {
            fetchAccountsList();
        }
        
        private Task SaveAttachmentForEveryMailItem(Folder folder)
        {
            Console.WriteLine("Saving attchments from " + folder.FolderPath);
            foreach(Object item in folder.Items)
            {
                if (item is MailItem)
                {
                    MailItem mailItem = ((MailItem)item);
                    if(mailItem.Attachments.Count > 0)
                    {
                        try
                        {
                            for (int i = 1; i <= mailItem.Attachments.Count; i++)
                            {
                                mailItem.Attachments[i].SaveAsFile(destination + mailItem.Attachments[i].Parent + "-" +  mailItem.Attachments[i].FileName);
                            }
                        }
                        catch(System.Runtime.InteropServices.COMException ex)
                        {
                            Console.WriteLine("Failure while trying get attachment from " + mailItem.Subject);
                            Console.WriteLine(ex.Message);
                        }
                        Console.WriteLine("Saving attachments from email " + mailItem.Subject);

                    }
                    
                }
            }
            return Task.CompletedTask;
        }

        private async Task<Folder> SearchFolder(string folderName, Folder root)
        {
            Folder found = null;
            if (root == null)
            {
                throw new System.Exception("Root folder is null");
            }
            found = await SearchFolderWorker(root, folderName);
            return found;
        }

        private async Task<Folder> SearchFolderWorker(Folder folder, String folderName)
        {
           Folders childFolders = folder.Folders;

            if (childFolders != null && childFolders.Count > 0)
            {
                foreach (Folder childFolder in childFolders)
                {
                    if (childFolder.Name == folderName)
                    {
                        return childFolder;
                    }
                    await SearchFolderWorker(childFolder, folderName);  
                }
            }
            return null;
        }


    }
}
