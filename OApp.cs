using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
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
        string[] selectedFolders;
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
                    throw new System.Exception("Account not found !");
            }
        }

        public string[] SelectedFolder
        {
            get
            {
                return selectedFolders;
            }
            set
            {
                selectedFolders = value;
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

        public string[] AccountsAvailable
        {
            get
            {
                return accounts.Keys.ToArray();
            }
        }

        public async Task<string[]> FetchFolderList()
        {
             var task = await Task.Run(() =>
            {
                List<string> foldersArray = new List<string>();
                EnumerateFolders(accounts[selectedAccount], (folder) =>
                {
                    foldersArray.Add(folder.FolderPath);
                });

                return foldersArray.ToArray();
            });
            return task;

        }

        public async Task SaveAttachments(string folderName)
        {
            
            Folder folder = await SearchFolder(folderName, accounts[selectedAccount]);
            if(folder == null)
            {
                throw new System.Exception("Folder " + folderName + " not found.");
            }
            else
            {
                await SaveAttachmentForEveryMailItem(folder);
                EnumerateFolders(folder,async (subFolder) =>
                {
                    await SaveAttachmentForEveryMailItem(folder);
                });
            }

        }

        public Task SaveAttachments()
        {
            List<Task> tasksList = new List<Task>(); 
            foreach(string folderName in selectedFolders)
            {
                var task = SaveAttachments(folderName);
                tasksList.Append(task);
                task.Start();
            }
            return Task.WhenAll(tasksList.ToArray());
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

                outlookApplication = new Application();
                outlookNamespace = outlookApplication.GetNamespace("MAPI");
                accounts = new Dictionary<string, Folder>();
                _initialization();
        }

        public void closeInstance()
        {
            if(outlookApplication != null)
            {
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
        
        private async Task SaveAttachmentForEveryMailItem(Folder folder)
        {
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
                                await mailItem.Attachments[i].SaveAsFile(destination + $"\\{folder.Name}\\" + mailItem.Attachments[i].Parent +  mailItem.Attachments[i].FileName);
                            }
                        }
                        catch(System.Runtime.InteropServices.COMException ex)
                        {
                            throw new  SystemException("Error while trying download attachment from email :  " + ex.Message + "\n" + "Issues with : " + mailItem.Subject);
                        }
                    }
                    
                }
            }
         
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
