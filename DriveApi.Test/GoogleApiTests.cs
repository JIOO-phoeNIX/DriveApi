using System.Web;
using Project_Api.GoogleDriveApi;
using Xunit;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Diagnostics;
using Google.Apis.Drive.v3.Data;
using Xunit.Abstractions;

namespace Project_Api.Test.GoogleDriveApiTests
{

    public class GoogleApiTests
    {
        private readonly ITestOutputHelper output;

        public GoogleApiTests(ITestOutputHelper output)
        {
            this.output = output;
        }

        /// <summary>
        /// Test the OAuth 2 Credentials process
        /// </summary>
        [Fact]
        public void GetUserCredentialTest()
        {
            // Arrange            

            // Act
            var actual = GoogleApi.GetUserCredential();

            // Assert
            Assert.NotNull(actual);
        }

        /// <summary>
        /// Test the service method used to upload the file
        /// </summary>
        [Fact]
        public void CreateServiceTest()
        {
            // Arrange

            // Act
            var actual = GoogleApi.CreateService();

            // Assert 
            Assert.NotNull(actual);
        }

        /// <summary>
        /// Test the upload method used to upload the file to Google Drive
        /// </summary>
        [Theory]
        [InlineData("GoogleDriveApiTestFiles/imagetest.jpg")]
        [InlineData("GoogleDriveApiTestFiles/bootstraptest.css")]
        [InlineData("GoogleDriveApiTestFiles/eng.pdf")]
        public void UploadFileTest(string filename)
        {
            // Arrange                      

            // Act           
            var actual = GoogleApi.UploadFile(filename);

            // Arrange
            Assert.NotNull(actual);
        }

        /// <summary>
        /// Test the method used to get the files from the root directory of Google Drive
        /// </summary>
        /// <param name="expectedFileName"></param>
        [Theory]
        [InlineData("GoogleDriveApiTestFiles/eng.pdf")]
        public void GetFilesInRootFolderTest(string expectedFileName)
        {
            // Arrange

            // Act
            List<Google.Apis.Drive.v3.Data.File> actual = GoogleApi.GetFilesInRootFolder();

            // Assert
            Assert.NotNull(actual);
            Assert.True(actual.Count > 0);
            //Assert.Equal(expectedFileName, actual[0].Name);
            foreach (var file in actual)
            {
                output.WriteLine(file.Id + " " + file.Name);
            }
        }

        /// <summary>
        /// Test the function used to get the files in a specific folder in Google drive 
        /// </summary>
        /// <param name="folderName">The name of the folder on Google drive</param>
        /// <param name="expectedLenght">The number of files in the folder</param>
        /// <param name="expectedFileName">The name of the first file in the folder</param>
        [Theory]
        [InlineData("FOLDER1", 3, "Workings.aspx")]
        [InlineData("Doc", 2, "doc.pdf")]
        [InlineData("Documents", 3, "InnerDocumentFolder")]
        [InlineData("NewEmptyFolder", 3, "InnerDocumentFolder")]
        public void GetFilesInSpecificFolderTest(string folderName, int expectedLenght, string expectedFileName)
        {
            // Arrange

            // Act
            List<Google.Apis.Drive.v3.Data.File> actual = GoogleApi.GetFilesInSpecificFolder(folderName);

            foreach (var file in actual)
            {
                Debug.WriteLine(file.Name);
            }

            // Assert
            Assert.NotNull(actual);
            Assert.Equal(expectedLenght, actual.Count);
            Assert.Equal(expectedFileName, actual[0].Name);
        }

        /// <summary>
        /// This is the test method for the function used to upload a file to a specific folder on Google drive
        /// </summary>
        /// <param name="folderName">The name of the folder</param>
        /// <param name="fileName">The name of the file</param>
        [Theory]
        [InlineData("Documents", "GoogleDriveApiTestFiles/imagetest.jpg")]
        [InlineData("NullFolder", "GoogleDriveApiTestFiles/bootstraptest.css")]
        public void UploadFileToSpecificFolderTest(string folderName, string fileName)
        {
            // Arrange

            // Act
            var actual = GoogleApi.UploadFileToSpecificFolder(folderName, fileName);

            // Assert
            Assert.NotNull(actual);
        }

        /// <summary>
        /// This is the test method for the method that is used to copy a file from one folder into another
        /// </summary>
        /// <param name="fromFolder">The folder to copy the file from</param>
        /// <param name="toFolder">The folder where the file will be copied into</param>
        /// <param name="fileName">The name of the file to be copied</param>
        [Theory]
        [InlineData("Folder1", "FolderTest", "Site.Master.vb")]
        public void CopyFileToNewLocationTest(string fromFolder, string toFolder, string fileName)
        {
            // Act
            var actual = GoogleApi.CopyFileToNewLocation(fromFolder, toFolder, fileName);

            // Assert
            Assert.NotNull(actual);
            Assert.True(actual.Result.Parents.Count > 1);
            Assert.Equal(fileName, actual.Result.Name);
        }

        /// <summary>
        /// This method is used to test the function that gets the metadata of a file using the file name
        /// </summary>
        /// <param name="fileName">The file</param>
        [Theory]
        [InlineData("style.css")]
        [InlineData("fillnot")]
        [InlineData("Site.Master.vb")]
        [InlineData("About.aspx")]
        public void GetFileMetadataUsingFileNameTest(string fileName)
        {
            // Act
            var actual = GoogleApi.GetFileMetadataUsingFileName(fileName);

            // Assert
            Assert.NotNull(actual);
            Assert.Equal(fileName, actual.FileName);
            Debug.WriteLine(actual.ToString());
        }

        /// <summary>
        /// This method is used to test the function that gets the metadata of a file using the file id
        /// </summary>
        /// <param name="fileId">The Id of the file</param>
        [Theory]
        [InlineData("16DxpeDJFmsPFuYXwcrSRjwAJFofkUBXj")]
        [InlineData("erwrrrrr")]
        [InlineData("1ICclh6t6wbCwnz9rNSndOHumM7uwVxhB")]
        public void GetFileMetadataUsingFileIdTest(string fileId)
        {
            // Act
            var actual = GoogleApi.GetFileMetadataUsingFileId(fileId);

            // Assert
            Assert.NotNull(actual);
            Debug.WriteLine(actual.ToString());
        }

        /// <summary>
        /// This is the method used to test the method that get file by content using the file Name as the input paramater
        /// </summary>
        /// <param name="fileName">The name of the file</param>
        [Theory]
        [InlineData("About.aspx")]
        public void GetFileByFileNameTest(string fileName)
        {
            // Act
            var actual = GoogleApi.GetFileByFileName(fileName);

            // Assert 
            Assert.NotNull(actual);
            Debug.WriteLine(actual);
        }

        /// <summary>
        /// This is the method used to test the method that get file by content using the file Id as the input parameter
        /// </summary>
        /// <param name="fileId">The Id of the file</param>
        [Theory]
        [InlineData("16DxpeDJFmsPFuYXwcrSRjwAJFofkUBXj")]
        [InlineData("erwrrrrr")]
        [InlineData("1ICclh6t6wbCwnz9rNSndOHumM7uwVxhB")]
        public void GetFileByFileIdTest(string fileId)
        {
            // Act
            var actual = GoogleApi.GetFileByFileId(fileId);

            // Assert 
            Assert.NotNull(actual);
            Debug.WriteLine(actual);
        }

        /// <summary>
        /// This method is used to test the method that is used to search for files
        /// </summary>
        /// <param name="searchText">The search parameter</param>
        [Theory]
        [InlineData("pdfcssaspx")]
        [InlineData("1ICclh6t6wbCwnz9rNSndOHumM7uwVxhB")]
        [InlineData("1ICclh6t6wbCwn")]
        public void SearchForFileTest(string searchText)
        {
            // Act 
            var actual = GoogleApi.SearchForFile(searchText);

            // Assert
            Assert.NotNull(actual);

            foreach (var file in actual)
            {
                Debug.WriteLine(file.Name);
                output.WriteLine(file.Name);
            }

        }


        [Theory]
        [InlineData("GoogleDriveApiTestFiles/Testdoc.txt", "15VAmgPspxgw8RDD684RqmaKauATTLePc")]
        public void UpdateFileTest(string newFileName, string fileId)
        {
            // Act
            var actual = GoogleApi.UpdateFile(newFileName, fileId);

            // Assert
            //Assert.NotNull(actual);
            //Assert.Equal(fileId, actual.Id);
            output.WriteLine(actual.Name, " ", actual.Id);

        }
    }
}
