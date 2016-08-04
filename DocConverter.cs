using DocumentFormat.OpenXml.Packaging;
using Microsoft.AspNet.Hosting;
using Microsoft.AspNet.Http;
using Microsoft.Net.Http.Headers;
using OpenXmlPowerTools;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using WebApplication_Webease_.Models;

//Create a folder in the webroot path, for the upload files.
//Get the file name of the uploaded file.
//Change the extension of the file name, to a razor view type.
//Create a new path for the razor view in the upload files folder.
//Create a new path for the uploaded file in the upload files folder.
//Save the path of the uploaded file, so has to be ableto use it later.
//Using OpenXml wordprocessingml Api
//Open the path of the uploaded file with the Api
//Set the html title of the razor view page.
//COnvert the page from word document to html, apply the page title.
//Delete the path of the uploaded file.
//Move the newly created razor view to the views folder of the hosting environment.
//Save the blogbody memebr of the blog class as the file name of the newly created razor file.

namespace WebApplication_Webease_.Services
{
    public class DocConverter : IDocConverter
    {
        private readonly IHostingEnvironment _environment;
      
        private static string fileName { get; set; }
        public DocConverter(IHostingEnvironment env)
        {
            _environment = env;
            
        }

        private static string ChangeFileExtension(string firstfileext, string secondfileext, IFormFile file)
        {
            var newDocExtension = "";
            var filename = ContentDispositionHeaderValue.Parse(file.ContentDisposition).FileName.Trim('"');
            if (fileName.EndsWith(firstfileext))
            {
                newDocExtension = fileName.Replace(firstfileext, secondfileext);
            }
            return newDocExtension;
        }

        private static string GetViewsDirectory(string parentPath)
        {
            var ViewdDir = "";
            var dirs = Directory.GetDirectories(parentPath);
            foreach (var dir in dirs)
            {
                if (dir.Contains("Views")) ViewdDir = dir;
            }
            return ViewdDir;
        }

        private static DirectoryInfo ReturnHomeDir(IEnumerable<DirectoryInfo> dirs)
        {
            DirectoryInfo homeDir = null;
            foreach (var dir in dirs)
            {
                if (dir.Name.Equals("Home"))
                    homeDir = dir;
            }
            return homeDir;
        }
        
        private static string  ConvertToRelativePath(string newlyCreatedRazorFullPath)
        {
            string ApprovedPath = null;
            var BlogFolder = "BlogUploads";
            var indexOfViewsFolder = newlyCreatedRazorFullPath.Contains(BlogFolder) ? newlyCreatedRazorFullPath.LastIndexOf(BlogFolder): -1;
            var relativePath = newlyCreatedRazorFullPath.Substring(indexOfViewsFolder);
            if (relativePath.EndsWith(".cshtml") && relativePath.Contains("\\"))
            {
                var delRazorExt = relativePath.Remove(relativePath.IndexOf(".cshtml"));
                ApprovedPath = delRazorExt.Replace("\\", "/");
            }

            return ApprovedPath;
        }

        public async Task ConvertWordToRazorViewAsync(IFormFile BlogBody,Blog Blog)
        {
            fileName = ContentDispositionHeaderValue.Parse(BlogBody.ContentDisposition).FileName.Trim('"');
            if (fileName.EndsWith(".docx"))
            {
                var oldBlogFileUploads = new DirectoryInfo($"{_environment.WebRootPath}\\BlogFiles");
                if (oldBlogFileUploads.Exists)
                {
                    if(oldBlogFileUploads.GetFiles() != null)
                        foreach (var file in oldBlogFileUploads.GetFiles())
                        {
                            file.Delete();
                        }
                    oldBlogFileUploads.Delete();
                }
                var BlogFiles = Directory.CreateDirectory("BlogFiles");
                var BlogFileUploads = Path.Combine(_environment.WebRootPath, BlogFiles.FullName);
                

                var newFile = fileName;
                var newHtmldoc = ChangeFileExtension(".docx", ".cshtml", BlogBody);


                var newHtmldocPath = Path.Combine(BlogFileUploads, newHtmldoc);
                

                var newPath = Path.Combine(BlogFileUploads, newFile);
       
                await BlogBody.SaveAsAsync(newPath);

                using (WordprocessingDocument doc = WordprocessingDocument.Open(newPath, true))
                {
                    
                    var settings = new HtmlConverterSettings
                    {
                        PageTitle = Blog.BlogTitle
                    };

                    var htmlDocument = HtmlConverter.ConvertToHtml(doc, settings);
                    File.WriteAllText(newHtmldocPath, htmlDocument.ToStringNewLineOnAttributes());

                    //get folder: "BlogFiles"
                    var BlogUploadsdir = new DirectoryInfo(newHtmldocPath).Parent.ToString();
                    
                    //get the parent folder full path, i.e the application name.
                    var parentDir = new DirectoryInfo(BlogUploadsdir).Parent.FullName.ToString();
                    
                    //get the "Views" folder of the application.
                    var viewsDir = GetViewsDirectory(parentDir);

                    //get a collection of directories under "Views" folder and create "BlogUploads" folder.
                    var ViewsDirs = new DirectoryInfo(viewsDir).EnumerateDirectories();
                    var blogUploadsDir = ReturnHomeDir(ViewsDirs).CreateSubdirectory("BlogUploads");

                    //create new path for the newly created Razor view file and move it there.
                    var viewHtmldocPath = Path.Combine(blogUploadsDir.FullName, newHtmldoc);
                    File.Move(newHtmldocPath, viewHtmldocPath); 

                    //persist the path to the newly created Razor View in the database.
                    Blog.BlogBody = $"{ConvertToRelativePath(viewHtmldocPath)}";

                }
            }
        }
    }
}
