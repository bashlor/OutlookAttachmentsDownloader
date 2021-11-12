using Microsoft.Office.Interop.Outlook;
using OutlookAttachmentsDownloader.Exceptions;
using System;
using System.Collections.Generic;
using System.IO;
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


        private string selectedAccount;

        private readonly string[] emails;

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

        public string[] SelectedFolders { get; set; }

        public string Destination { get; set; }

        public string[] AccountsAvailable
        {
            get
            {
                return emails;
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
            
            Folder folder = SearchFolder(folderName, accounts[selectedAccount]);
            if(folder == null)
            {
                throw new OAppException("Folder " + folderName + " not found.");
            }
            else
            {
                await Task.Run(()=> {
                    SaveAttachmentForEveryMailItem(folder);
                });
                EnumerateFolders(folder,async (subFolder) =>
                {
                    await Task.Run(() => { SaveAttachmentForEveryMailItem(folder); });
                });
            }

        }

        public async Task SaveAttachments()
        {
            foreach(string folderName in SelectedFolders)
            {
                await SaveAttachments(folderName);
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



        private void FetchAccountsList()
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
                SelectedFolders = Array.Empty<string>();
                Destination = "";
                Initialization();
                emails = accounts.Keys.ToArray();
        }

        public void CloseInstance()
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

        private  void Initialization()
        {
            FetchAccountsList();
        }
        
        private  void SaveAttachmentForEveryMailItem(Folder folder)
        {
            
            foreach(Object item in folder.Items)
            {
                string finalPath = Path.Combine(Destination,folder.Name);
                if (!Directory.Exists(finalPath))
                    Directory.CreateDirectory(finalPath);
                if (item is MailItem)
                {
                    MailItem mailItem = ((MailItem)item);
                    if(mailItem.Attachments.Count > 0)
                    {
                        try
                        {
                            for (int i = 1; i <= mailItem.Attachments.Count; i++)
                            {
                                 mailItem.Attachments[i].SaveAsFile(Path.Combine(finalPath,mailItem.Attachments[i].FileName));
                                Console.WriteLine("Downloaded : " + mailItem.Attachments[i].FileName);
                            }
                        }
                        catch(System.Runtime.InteropServices.COMException ex)
                        {
                            Console.WriteLine("Error while trying download attachment from email :  " + ex.Message + "\n" + "Issues with : " + mailItem.Subject);
                        }
                    }
                    
                }
            }
         
        }

        private Folder SearchFolder(string folderPath, Folder root)
        {
              
            List<string> splittedPath = folderPath.Split(Path.DirectorySeparatorChar).ToList();
            splittedPath.RemoveAll(pathElement => pathElement.Length == 0);
            splittedPath.RemoveAt(0);
            Folder found = null;
            if (root == null)
            {
                throw new OAppException("Root folder is null");
            }
            found = SearchFolderWorker(root,splittedPath);
            return found;
        }

        private Folder SearchFolderWorker(Folder folder, List<string> folderPath,bool _continue=true,int level = 0)
        {
           Folders childFolders = folder.Folders;
            Folder result = null;
            if (childFolders != null && childFolders.Count > 0)
            {
                if (_continue)
                {
                    foreach (Folder childFolder in childFolders)
                    {
                        if (level < folderPath.Count)
                        {
                            if (childFolder.Name == folderPath[level] && level + 1 == folderPath.Count)
                            {
                                result = childFolder;
                                break;

                            }
                            else if (childFolder.Name == folderPath[level] && level + 1 < folderPath.Count)
                            {
                                return SearchFolderWorker(childFolder, folderPath, _continue, level + 1);
                                
                            }
                        }
                    }
                    return result;
                }
                
                        
            }
            return null;
        }
    }
}
