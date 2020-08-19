using Google;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Download;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web;

namespace Project_Api.GoogleDriveApi
{
    public class GoogleApi
    {
        private static readonly string[] scopes = { DriveService.Scope.Drive }; //define the scopes
        private static readonly string applicationName = "Assignment";


        /// <summary>
        /// This method is used to get the user's OAuth 2.0 credentials
        /// </summary>
        /// <returns>The user's OAuth 2.0 credentials</returns>
        public static UserCredential GetUserCredential()
        {
            UserCredential credential;

            //Read the Client Credentials from the client_secret.json file and use it to get the 
            //OAuth 2.0 credentials and store it in the folder C:\users_credentials
            using (FileStream stream = new FileStream(@"C:\client_secret.json", FileMode.Open, FileAccess.Read))
            {
                //Set the folder path where the OAuth 2.0 credentials is stored
                string folderPath = @"C:\";
                string credentialPath = Path.Combine(folderPath, "users_credentials");

                //Get the OAuth 2.0 credentials after authenticating the user and user gives access
                //and store it in the folder C:\users_credentials
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credentialPath, true)).Result;
            }

            return credential;
        }

        /// <summary>
        /// This method creates a Drive API service using the user's OAuth 2.0 credentials
        /// </summary>
        /// <returns>Google Drive API service</returns>
        public static DriveService CreateService()
        {
            // Create a Drive API service.
            DriveService service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = GetUserCredential(),
                ApplicationName = applicationName,
            });

            return service;
        }

        /// <summary>
        /// This method trys to upload the file to Google Drive using the Drive API Service       
        /// </summary>
        /// <param name="file">The file to be uploaded</param>
        public static FilesResource.CreateMediaUpload UploadFile(string file)
        {
            DriveService service = CreateService();

            var fileData = new Google.Apis.Drive.v3.Data.File
            {
                Name = file,
                MimeType = MimeMapping.GetMimeMapping(file)
            };

            //Create a media upload that is used to upload the file to Google Drive
            FilesResource.CreateMediaUpload mediaUpload;

            //Open the file to be uploaded to Google Drive
            using (FileStream stream = new FileStream(file, FileMode.Open))
            {
                //Create the file
                mediaUpload = service.Files.Create(fileData, stream, fileData.MimeType);
                mediaUpload.Fields = "id";

                //Upload the file to Google Drive
                mediaUpload.UploadAsync();
            }

            return mediaUpload;

        }

        /// <summary>
        /// This method is used to get all the files in the root directory of Google Drive
        /// </summary>
        /// <returns>The List of files</returns>
        public static List<Google.Apis.Drive.v3.Data.File> GetFilesInRootFolder()
        {
            try
            {
                DriveService service = CreateService();

                // define parameters of request.
                FilesResource.ListRequest fileListRequest = service.Files.List();
                fileListRequest.Fields = "nextPageToken, files(*)";

                //get file list.
                IList<Google.Apis.Drive.v3.Data.File> files = fileListRequest.Execute().Files;
                List<Google.Apis.Drive.v3.Data.File> fileList = new List<Google.Apis.Drive.v3.Data.File>();

                if (files != null && files.Count > 0)
                {
                    foreach (var file in files)
                    {
                        Google.Apis.Drive.v3.Data.File driveFile = new Google.Apis.Drive.v3.Data.File
                        {
                            Id = file.Id,
                            Name = file.Name,
                            Size = file.Size,
                            CreatedTime = file.CreatedTime
                        };

                        fileList.Add(driveFile);
                    }
                }
                return fileList;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }

        /// <summary>
        /// This method is used to get all the files in a particular folder in Google Drive
        /// </summary>
        /// <param name="folderName">The name of the folder</param>
        /// <returns>All the files in the folder</returns>
        public static List<Google.Apis.Drive.v3.Data.File> GetFilesInSpecificFolder(string folderName)
        {
            DriveService service = CreateService();

            // define parameters of request.
            FilesResource.ListRequest fileListRequest = service.Files.List();

            //listRequest.PageToken = 10;
            fileListRequest.Fields = "nextPageToken, files(*)";

            //get file in the root folder.
            IList<Google.Apis.Drive.v3.Data.File> rootFolderFiles = fileListRequest.Execute().Files;

            string folderId = null;

            //Get the folderId of the folderName passed as parameter
            if (rootFolderFiles != null && rootFolderFiles.Count > 0)
            {
                foreach (var file in rootFolderFiles)
                {
                    if (file.Name.ToLower() == folderName.ToLower())
                    {
                        folderId = file.Id;
                        break;
                    }
                }
;
            }

            //Thrown an exception if the folderId is null indicating that the folderName doesn't exist
            if (folderId == null)
            {
                throw new GoogleApiException("Folder does not exist");
            }

            List<Google.Apis.Drive.v3.Data.File> fileList = new List<Google.Apis.Drive.v3.Data.File>();

            //Get all the files in the folder using the folderId and store them in the fileList array
            if (rootFolderFiles != null && rootFolderFiles.Count > 0)
            {
                foreach (var file in rootFolderFiles)
                {
                    foreach(var parent in file.Parents)
                    {
                        if (parent == folderId)
                        {
                            Google.Apis.Drive.v3.Data.File driveFile = new Google.Apis.Drive.v3.Data.File
                            {
                                Id = file.Id,
                                Name = file.Name,
                                Size = file.Size,
                                CreatedTime = file.CreatedTime
                            };

                            fileList.Add(driveFile);
                        }
                    }                   
                }
            }

            // Thrown exception if List is empty
            if (fileList.Count == 0)
            {
                throw new GoogleApiException("Folder doesn't contain any file.");
            }
            return fileList;

        }

        /// <summary>
        /// This method is used to upload a file to a specific folder in Google drive
        /// </summary>
        /// <param name="folderName">The nameof the folder in Google Drive</param>
        /// <param name="fileName">The name of the file to upload</param>
        /// <returns></returns>
        public static FilesResource.CreateMediaUpload UploadFileToSpecificFolder(string folderName, string fileName)
        {
            DriveService service = CreateService();

            // define parameters of request.
            FilesResource.ListRequest fileListRequest = service.Files.List();

            //listRequest.PageToken = 10;
            fileListRequest.Fields = "nextPageToken, files(*)";

            //get file in the root folder.
            IList<Google.Apis.Drive.v3.Data.File> rootFolderFiles = fileListRequest.Execute().Files;

            string folderId = null;

            //Get the folderId of the folderName passed as parameter
            if (rootFolderFiles != null && rootFolderFiles.Count > 0)
            {
                foreach (var file in rootFolderFiles)
                {
                    if (file.Name.ToLower() == folderName.ToLower())
                    {
                        folderId = file.Id;
                        break;
                    }
                }
            }

            //Thrown an exception if the folderId is null indicating that the folderName doesn't exist
            if (folderId == null)
            {
                throw new GoogleApiException("Folder does not exist");
            }

            var fileData = new Google.Apis.Drive.v3.Data.File
            {
                Name = fileName,
                MimeType = MimeMapping.GetMimeMapping(fileName),
                Parents = new List<string> { folderId }
            };

            //Create a media upload that is used to upload the file to Google Drive
            FilesResource.CreateMediaUpload mediaUpload;

            //Open the file to be uploaded to Google Drive
            using (FileStream stream = new FileStream(fileName, FileMode.Open))
            {
                //Create the file
                mediaUpload = service.Files.Create(fileData, stream, fileData.MimeType);
                mediaUpload.Fields = "id";

                //Upload the file to Google Drive
                mediaUpload.UploadAsync();
            }

            return mediaUpload;
        }

        /// <summary>
        /// This is the method that can be used to copy a file from one folder into another in folder
        /// in Google Drive.
        /// </summary>
        /// <param name="fromFolder">The folder to copy from</param>
        /// <param name="toFolder">The folder to copy into</param>
        /// <param name="fileName">The name of the file to be copied</param>
        /// <returns>The copied file</returns>
        public static async Task<Google.Apis.Drive.v3.Data.File> CopyFileToNewLocation(string fromFolder, string toFolder, string fileName)
        {
            DriveService service = CreateService();

            string fileId = null, fromFolderId = null, toFolderId = null;
            Google.Apis.Drive.v3.Data.File filedata;

            // Get all the files in the Google Drive            
            FilesResource.ListRequest fileListRequest = service.Files.List();
            fileListRequest.Fields = "nextPageToken, files(*)";

            //get file in the root folder.
            IList<Google.Apis.Drive.v3.Data.File> rootFolderFiles = fileListRequest.Execute().Files;

            //Get the Id of the fromfolder
            if (rootFolderFiles != null && rootFolderFiles.Count > 0)
            {
                foreach (var file in rootFolderFiles)
                {
                    if (file.Name.ToLower() == fromFolder.ToLower())
                    {
                        fromFolderId = file.Id;
                        break;
                    }
                }
            }

            //Throw an exception if the fromFolderId is still null indicating that the from folder doesn't exist
            if (fromFolderId == null)
            {
                throw new GoogleApiException("Source folder doesn't exist");
            }

            //Get te Id of the toFolder
            if (rootFolderFiles != null && rootFolderFiles.Count > 0)
            {
                foreach (var file in rootFolderFiles)
                {
                    if (file.Name.ToLower() == toFolder.ToLower())
                    {
                        toFolderId = file.Id;
                        break;
                    }
                }
            }

            //Throw an exception if the toFolderId is still null indicating that the to folder doesn't exist
            if (toFolderId == null)
            {
                throw new GoogleApiException("Destination folder doesn't exist");
            }

            //Get the Id and meatadata of the file in the fromFolder
            if (rootFolderFiles != null && rootFolderFiles.Count > 0)
            {
                foreach (var file in rootFolderFiles)
                {
                    foreach (var parent in file.Parents)
                    {
                        if (parent == fromFolderId && file.Name == fileName)
                        {
                            fileId = file.Id;
                            filedata = file;
                            break;
                        }
                    }
                }
            }

            //Throw an exception to indicate that the file doesn't exist in the from folder
            if (fileId == null)
            {
                throw new GoogleApiException("File does not exist in the folder");
            }

            //copy the file to a new location
            FilesResource.UpdateRequest updateRequest = service.Files.Update(new Google.Apis.Drive.v3.Data.File(), fileId);
            updateRequest.Fields = "*";
            updateRequest.AddParents = toFolderId;

            //Execute the request
            filedata = await updateRequest.ExecuteAsync();

            return filedata;
        }

        /// <summary>
        /// This function is used to return the metadata of a file 
        /// </summary>
        /// <param name="fileName">The file name</param>
        /// <returns>A class that holds the metadata of the file</returns>
        public static FileMetadata GetFileMetadataUsingFileName(string fileName)
        {
            DriveService service = CreateService();
            List<string> parentFolderId = new List<string>();

            // define parameters of request.
            FilesResource.ListRequest fileListRequest = service.Files.List();
            fileListRequest.Fields = "nextPageToken, files(*)";

            //get file in the root folder.
            IList<Google.Apis.Drive.v3.Data.File> rootFolderFiles = fileListRequest.Execute().Files;
            Google.Apis.Drive.v3.Data.File googleFile = null;

            //get the file from the store if it exist           
            if (rootFolderFiles != null && rootFolderFiles.Count > 0)
            {
                foreach (var file in rootFolderFiles)
                {
                    if (file.Name.ToLower() == fileName.ToLower())
                    {
                        googleFile = file;
                        break;
                    }
                }
            }

            //throw exception if the file wasn't found
            if (googleFile == null)
            {
                throw new GoogleApiException("File doesn't exist");
            }

            //Create a FileMetadata and populate it with the Google Drive file metadata
            FileMetadata fileMetadata = new FileMetadata
            {
                FileName = googleFile.Name,
                FileSize = googleFile.Size,
                LastModifiedDate = googleFile.ModifiedTime,
                CreatedDate = googleFile.CreatedTime
            };

            // Get all the parent folder Id
            foreach (var parent in googleFile.Parents)
            {
                parentFolderId.Add(parent);
            }

            //Get all the names of the folder the file belongs to using the parent id
            foreach (var parentId in parentFolderId)
            {
                foreach (var file in rootFolderFiles)
                {
                    if (file.Id == parentId)
                    {
                        fileMetadata.ParentFolderName.Add(file.Name);
                    }
                }
            }

            return fileMetadata;
        }

        /// <summary>
        /// This function is used to return the metadata of a file using the file id as the input parameter
        /// </summary>
        /// <param name="fileId">The Id of the file</param>
        /// <returns>A class that holds the metadata of the file</returns>
        public static FileMetadata GetFileMetadataUsingFileId(string fileId)
        {
            DriveService service = CreateService();
            List<string> parentFolderId = new List<string>();

            // define parameters of request.
            FilesResource.ListRequest fileListRequest = service.Files.List();
            fileListRequest.Fields = "nextPageToken, files(*)";

            //get file in the root folder.
            IList<Google.Apis.Drive.v3.Data.File> rootFolderFiles = fileListRequest.Execute().Files;
            Google.Apis.Drive.v3.Data.File googleFile = null;

            //get the file from the store if it exist           
            if (rootFolderFiles != null && rootFolderFiles.Count > 0)
            {
                foreach (var file in rootFolderFiles)
                {
                    if (file.Id == fileId)
                    {
                        googleFile = file;
                        break;
                    }
                }
            }

            //throw exception if the file wasn't found
            if (googleFile == null)
            {
                throw new GoogleApiException("File doesn't exist");
            }

            //Create a FileMetadata and populate it with the Google Drive file metadata
            FileMetadata fileMetadata = new FileMetadata
            {
                FileName = googleFile.Name,
                FileSize = googleFile.Size,
                LastModifiedDate = googleFile.ModifiedTime,
                CreatedDate = googleFile.CreatedTime
            };

            // Get all the parent folder Id
            foreach (var parent in googleFile.Parents)
            {
                parentFolderId.Add(parent);
            }

            //Get all the names of the folder the file belongs to using the parent id
            foreach (var parentId in parentFolderId)
            {
                foreach (var file in rootFolderFiles)
                {
                    if (file.Id == parentId)
                    {
                        fileMetadata.ParentFolderName.Add(file.Name);
                    }
                }
            }

            return fileMetadata;
        }

        /// <summary>
        /// This function is used to download a file taking the file name as the input parameter.
        /// </summary>
        /// <param name="fileName">The name of the file</param>
        /// <param name="downloadPath">The path where the file is been downloaded into</param>
        /// <returns>A string confirmation that the file downloaded</returns>
        public static string GetFileByFileName(string fileName, string downloadPath = @"C:\Users\JIOO\source\repos\work\ApiProject\Project-Api\Project-Api.Test\DownloadedFiles")
        {
            string status = null;
            DriveService service = CreateService();
            Google.Apis.Drive.v3.Data.File googleFile = null;

            // define parameters of request.
            FilesResource.ListRequest fileListRequest = service.Files.List();
            fileListRequest.Fields = "nextPageToken, files(*)";

            //get file in the root folder.
            IList<Google.Apis.Drive.v3.Data.File> rootFolderFiles = fileListRequest.Execute().Files;

            //get the file from the store if it exist           
            if (rootFolderFiles != null && rootFolderFiles.Count > 0)
            {
                foreach (var file in rootFolderFiles)
                {
                    if (file.Name.ToLower() == fileName.ToLower())
                    {
                        googleFile = file;
                        break;
                    }
                }
            }

            //throw exception if the file wasn't found
            if (googleFile == null)
            {
                throw new GoogleApiException("File doesn't exist");
            }

            string filePath = Path.Combine(downloadPath, fileName);
            var request = service.Files.Get(googleFile.Id);
            var stream = new MemoryStream();

            request.MediaDownloader.ProgressChanged += (IDownloadProgress progress) =>
            {
                switch (progress.Status)
                {
                    case DownloadStatus.Downloading:
                        {
                            //Wait for download to complete
                            break;
                        }
                    case DownloadStatus.Completed:
                        {
                            status = "Download complete";
                            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                            {
                                stream.WriteTo(file);
                            }
                            break;
                        }
                    case DownloadStatus.Failed:
                        {
                            status = "Download failed";
                            break;
                        }

                    default:
                        break;
                }
            };
            request.DownloadAsync(stream);

            return status + " " + stream.ToString();
        }

        /// <summary>
        /// This function is used to download a file taking the file Id as the input parameter.
        /// </summary>
        /// <param name="fileId">The Id of the file</param>
        /// <param name="downloadPath">The download parameter</param>
        /// <returns></returns>
        public static string GetFileByFileId(string fileId, string downloadPath = @"C:\Users\JIOO\source\repos\work\ApiProject\Project-Api\Project-Api.Test\DownloadedFiles")
        {
            string status = null;
            DriveService service = CreateService();
            Google.Apis.Drive.v3.Data.File googleFile = null;

            // define parameters of request.
            FilesResource.ListRequest fileListRequest = service.Files.List();
            fileListRequest.Fields = "nextPageToken, files(*)";

            //get file in the root folder.
            IList<Google.Apis.Drive.v3.Data.File> rootFolderFiles = fileListRequest.Execute().Files;

            //get the file from the store if it exist           
            if (rootFolderFiles != null && rootFolderFiles.Count > 0)
            {
                foreach (var file in rootFolderFiles)
                {
                    if (file.Id == fileId)
                    {
                        googleFile = file;
                        break;
                    }
                }
            }

            //throw exception if the file wasn't found
            if (googleFile == null)
            {
                throw new GoogleApiException("File doesn't exist");
            }

            string filePath = Path.Combine(downloadPath, googleFile.Name);
            var request = service.Files.Get(googleFile.Id);
            var stream = new MemoryStream();

            request.MediaDownloader.ProgressChanged += (IDownloadProgress progress) =>
            {
                switch (progress.Status)
                {
                    case DownloadStatus.Downloading:
                        {
                            //Wait for download to complete
                            break;
                        }
                    case DownloadStatus.Completed:
                        {
                            status = "Download complete";
                            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
                            {
                                stream.WriteTo(file);
                            }
                            break;
                        }
                    case DownloadStatus.Failed:
                        {
                            status = "Download failed";
                            break;
                        }

                    default:
                        break;
                }
            };
            request.DownloadAsync(stream);

            return status + " " + stream.ToString();
        }

        /// <summary>
        /// This method is used to search for a file
        /// </summary>
        /// <param name="searchString">The search parameter</param>
        /// <returns>The files</returns>
        public static List<Google.Apis.Drive.v3.Data.File> SearchForFile(string searchString)
        {
            DriveService service = CreateService();

            // define parameters of request.
            FilesResource.ListRequest fileListRequest = service.Files.List();
            fileListRequest.Fields = "nextPageToken, files(*)";

            //Get the Google drive files.
            IList<Google.Apis.Drive.v3.Data.File> files = fileListRequest.Execute().Files;
            List<Google.Apis.Drive.v3.Data.File> searchResult = new List<Google.Apis.Drive.v3.Data.File>();

            if (files != null && files.Count > 0)
            {
                foreach (var file in files)
                {
                    if (searchString.ToLower().Contains(file.Name.ToLower()) || searchString.ToLower().Contains(file.Id.ToLower())
                        || searchString.ToLower().Contains(file.CreatedTimeRaw.ToLower()) || searchString.ToLower().Contains(file.MimeType.ToLower())
                        || searchString.ToLower().Contains((file.FileExtension?.ToLower() ?? "null")))
                    {
                        Google.Apis.Drive.v3.Data.File driveFile = new Google.Apis.Drive.v3.Data.File
                        {
                            Id = file.Id,
                            Name = file.Name,
                            Size = file.Size,
                            CreatedTime = file.CreatedTime
                        };

                        searchResult.Add(driveFile);
                    }
                }
            }

            // Thrown exception if List is empty
            if(searchResult.Count == 0)
            {
                throw new GoogleApiException("Search doesn't return any result.");
            }

            return searchResult;
        }

        /// <summary>
        /// This function is used to update a file metadata and content
        /// </summary>
        /// <param name="newFileName">The name of the new file</param>
        /// <param name="fileId">The Id of the file to update</param>
        /// <returns>The updated file</returns>
        public static Google.Apis.Drive.v3.Data.File UpdateFile(string newFileName, string fileId)
        {
            DriveService service = CreateService();
            Google.Apis.Drive.v3.Data.File googleFile = null;

            var request = service.Files?.Get(fileId) ?? throw new GoogleApiException("File doesn't exist");
            request.Fields = "*";

            googleFile = request.Execute();

            //throw exception if the file wasn't found
            if (googleFile == null)
            {
                throw new GoogleApiException("File doesn't exist");
            }

            //update the metadata of the Google file
            googleFile.Name = newFileName;
            googleFile.MimeType = newFileName;

            //Read the content of the new file
            byte[] byteArray = File.ReadAllBytes(newFileName);
            var stream = new MemoryStream(byteArray);

            //Update the content of the Google file with the content of the new file
            using (stream)
            {
                var updateFile = service.Files.Update(googleFile, googleFile.Id, stream, newFileName);

                //Execute the request               
                updateFile.UploadAsync();

                return updateFile.ResponseBody;
            }                             
        }
    }

    /// <summary>
    /// This class holds all the exception that can be thrown by the GoogleApiDrive
    /// </summary>
    public class GoogleApiException : ApplicationException
    {
        /// <summary>
        /// Create an execption using the message
        /// </summary>
        /// <param name="message"></param>
        public GoogleApiException(string message) : base(message)
        {

        }
    }

    /// <summary>
    /// This class is used to store all the needed metadata of a Google Drive file
    /// </summary>
    public class FileMetadata
    {
        public string FileName { get; set; }
        public long? FileSize { get; set; }
        public List<string> ParentFolderName { get; set; } = new List<string>();
        public DateTime? LastModifiedDate { get; set; }
        public DateTime? CreatedDate { get; set; }

        public override string ToString()
        {
            string parentNames = null;
            if (ParentFolderName.Count == 0)
            {
                parentNames = "My root folder";
            }
            else
            {
                foreach (var parentName in ParentFolderName)
                {
                    parentNames += parentName + ",\n";
                }
            }

            return $"File Name = {FileName}, \n File Size in bytes = {FileSize},\n, Created Date = {CreatedDate},\n" +
                $"Last Modified Date = {LastModifiedDate}\n, Parent Folder Name(s) = {parentNames}";
        }
    }
}
