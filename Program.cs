/*
 * 由SharpDevelop创建。
 * 用户： CaptGyrfalcon
 * 日期: 2022/10/2
 * 时间: 15:28

 */
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Net;
using System.Web;

using System.Reflection;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;

using Picture = DocumentFormat.OpenXml.Wordprocessing.Picture;

//Aspose要付费
//Open XML不要，但是使用有点麻烦

namespace FileAutoSort
{
    class Program
    {
        public static async Task Main(string[] args)
        {
            //请求Access Token
            var client = new HttpClient();
            var content = new StringContent("{}");
            client.DefaultRequestHeaders.Accept.Clear();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            //client.DefaultRequestHeaders.Add("Content-Type", "application/json");
            //不要乱动
            string clientID;
            string clientSecret;
            clientID = Console.ReadLine();
            clientSecret = Console.ReadLine();
            //考虑到实际情况，发布的时候记得把clientID和clientSecret直接带上...
            var response = await client.PostAsync("https://aip.baidubce.com/oauth/2.0/token?client_id=" + clientID + "&client_secret=" + clientSecret + "&grant_type=client_credentials", content);
            var responseString = await response.Content.ReadAsStringAsync();
            var jsonNode = JsonSerializer.Deserialize<JsonObject>(responseString);

            //access token
            string token = jsonNode["access_token"].ToString();
            Console.WriteLine("Access Token，请勿将其泄露给他人");
            Console.WriteLine(token);







            string currentFolderPath = System.Environment.CurrentDirectory;


            //考虑到不同设备的文件命名格式不同，需要对部分文件进行重命名
            string[] beforeFiles = Directory.GetFiles(currentFolderPath, "*.jpg");
            foreach (string beforeFile in beforeFiles)
            {

                string beforeName = Path.GetFileNameWithoutExtension(beforeFile);
                if (beforeName.Contains("_"))
                {
                    continue;
                }
                string timeString = beforeName.Split("IMG")[1];
                string time1 = timeString.Substring(0, 8);
                string time2 = timeString.Substring(8);
                string afterName = "IMG_" + time1 + "_" + time2 + ".jpg";
                File.Move(beforeFile, Path.Combine(currentFolderPath, afterName));
            }

            beforeFiles = Directory.GetFiles(currentFolderPath, "*.mp4");
            foreach (string beforeFile in beforeFiles)
            {

                string beforeName = Path.GetFileNameWithoutExtension(beforeFile);
                if (beforeName.Contains("_"))
                {
                    continue;
                }
                string timeString = beforeName.Split("VID")[1];
                string time1 = timeString.Substring(0, 8);
                string time2 = timeString.Substring(8);
                string afterName = "VID_" + time1 + "_" + time2 + ".mp4";
                File.Move(beforeFile, Path.Combine(currentFolderPath, afterName));
            }



            string[] files = Directory.GetFiles(currentFolderPath);
            List<FileInformation> fileList = new List<FileInformation>();




            foreach (string file in files)
            {
                if (IsData(file))
                {
                    fileList.Add(new FileInformation(file));
                }

            }
            fileList.Sort(new Comparison<FileInformation>(FileInformation.CompareFileDate));
            int index = 0;
            string currentFolderInside = "";
            foreach (FileInformation file in fileList)
            {
                if (file.fileName.EndsWith("mp4"))
                {
                    index++;
                    currentFolderInside = Path.Combine(currentFolderPath, index.ToString());
                    Directory.CreateDirectory(currentFolderInside);
                }
                File.Move(file.fileName, Path.Combine(currentFolderInside, file.GetFileName()));
            }


            string[] newDirectories = Directory.GetDirectories(Environment.CurrentDirectory);
            List<string> existingDirectories = new List<string>();
            string outputLog = "";
            foreach (string directory in newDirectories)
            {
                existingDirectories.Add(Path.GetFileNameWithoutExtension(directory));
            }
            int currentIndex = 1;

            //task begin
            List<Task> tasks = new List<Task>();



            foreach (string directory in newDirectories)
            {
                await Task.Delay(TimeSpan.FromMilliseconds(10));
                Console.WriteLine("开始任务 " + currentIndex.ToString());
                int taskId = currentIndex;
                currentIndex++;
                tasks.Add(
                Task.Run(() =>
                {
                    if (!int.TryParse(Path.GetFileNameWithoutExtension(directory).Substring(0, 1), out _))
                    {
                        return;
                    }
                    string targetCardImage = "";
                    string targetVideo = "";
                    string[] newFiles = Directory.GetFiles(directory, "*.jpg");

                    if (newFiles.Length != 3)
                    {
                        outputLog += "序号 " + Path.GetFileNameWithoutExtension(directory) + " 文件数量有误，请人工核实。\n";
                        return;
                    }

                    if (newFiles.Length > 0)
                    {
                        string oldest = newFiles[0];
                        FileInfo oldestInfo = new FileInfo(oldest);
                        foreach (string file in newFiles)
                        {
                            FileInfo fileInfo = new FileInfo(file);
                            if (string.Compare(file, oldest) < 0)
                            {
                                oldest = file;
                                oldestInfo = fileInfo;
                            }
                        }
                        targetCardImage = oldest;
                    }

                    string[] mpFile = Directory.GetFiles(directory, "*.mp4");
                    if (mpFile.Length > 0)
                    {
                        targetVideo = mpFile[0];
                    }
                    if (!string.IsNullOrEmpty(targetCardImage))
                    {
                        Console.WriteLine("Acquire Image Data: " + targetCardImage);
                        string rawJsonData = GetCardInformation(targetCardImage, token);
                        Console.WriteLine("Done with acquire Image Data: " + targetCardImage);
                        if (IsOK(rawJsonData))
                        {
                            string name = GetName(rawJsonData);
                            while (true)
                            {
                                if (existingDirectories.Contains(name))
                                {

                                    outputLog += name + "存在重名，请人工复核。\n";
                                    name += "a";
                                }
                                else
                                {
                                    break;
                                }
                            }
                            existingDirectories.Add(name);
                            string id = GetNumber(rawJsonData);
                            string parentDirectory = Directory.GetParent(directory).FullName;
                            string newDirectory = Path.Combine(parentDirectory, name);
                            Directory.Move(directory, newDirectory);
                            string[] movedFiles = Directory.GetFiles(newDirectory, "*.jpg");
                            string nearFile = "";
                            string farFile = "";
                            if (movedFiles.Length > 2)
                            {
                                Array.Sort(movedFiles);
                                nearFile = movedFiles[1];
                                farFile = movedFiles[2];
                            }
                            File.Move(movedFiles[0], Path.Combine(newDirectory, name + ".jpg"));
                            string[] movedVideos = Directory.GetFiles(newDirectory, "*.mp4");
                            if (movedVideos.Length > 0)
                            {
                                string videoPath = movedVideos[0];
                                File.Move(videoPath, Path.Combine(newDirectory, name + ".mp4"));
                            }
                            string documentPath = Path.Combine(newDirectory, name + ".docx");
                            File.Copy(Path.Combine(Environment.CurrentDirectory, "model.docx"), documentPath, true);
                            //Document doc = new Document(documentPath);
                            //doc.Range.Replace("$idcard", id);
                            //doc.Range.Replace("$name", name);
                            //NodeCollection nodes = doc.GetChildNodes(NodeType.Shape, true);
                            //for (int i = 0; i < 2; i++)
                            //{
                            //    Shape shape = (Shape)nodes[i];
                            //    string imagename = "";
                            //    if (i == 0)
                            //    {
                            //        imagename = nearFile;
                            //    }
                            //    else
                            //    {
                            //        imagename = farFile;
                            //    }
                            //    shape.ImageData.SetImage(imagename);
                            //}
                            //doc.Save(documentPath);

                            WordprocessingDocument doc = WordprocessingDocument.Open(documentPath, true);
                            var body = doc.MainDocumentPart.Document.Body;

                            foreach (var text in body.Descendants<Text>())
                            {
                                if (text.Text.Contains("$IDCard"))
                                {
                                    text.Text = text.Text.Replace("$IDCard", id);
                                }
                                if (text.Text.Contains("$Name"))
                                {
                                    text.Text = text.Text.Replace("$Name", name);
                                }
                            }
                            //覆写近景远景图片
                            ImagePart imagePart = (ImagePart)doc.MainDocumentPart.GetPartById("rId4");
                            byte[] imageBytes = File.ReadAllBytes(nearFile);
                            BinaryWriter writer = new BinaryWriter(imagePart.GetStream());
                            writer.Write(imageBytes);
                            writer.Close();

                            ImagePart imagePart2 = (ImagePart)doc.MainDocumentPart.GetPartById("rId5");
                            byte[] imageBytes2 = File.ReadAllBytes(farFile);
                            BinaryWriter writer2 = new BinaryWriter(imagePart2.GetStream());
                            writer2.Write(imageBytes2);
                            writer2.Close();

                            doc.Save();
                            doc.Dispose();

                        }
                        else
                        {
                            outputLog += "序号 " + Path.GetFileNameWithoutExtension(directory) + " 身份证图片存在问题，请人工复核。\n";
                        }
                    }
                }));

            }
            await Task.WhenAll(tasks);
            //task end




            if (string.IsNullOrWhiteSpace(outputLog) || string.IsNullOrEmpty(outputLog))
            {
                outputLog = "不存在异常情况（如果存在代签，请人工挑拣复核）";
            }
            else
            {
                outputLog += "\n" + "存在上述异常情况（如果存在代签，请人工挑拣复核）";
            }
            File.WriteAllText(Path.Combine(currentFolderPath, "处理结果.txt"), outputLog);
            Console.Clear();
            Console.WriteLine(outputLog);
            Console.WriteLine("处理全部完成，按回车三次退出程序。");
            Console.ReadLine();
            Console.ReadLine();
            Console.ReadLine();


        }
        public static bool IsData(string s)
        {
            return s.EndsWith("mp4") || s.EndsWith("jpg");
        }

        //调用百度智能云OCR API进行处理
        public static string GetCardInformation(string filePath, string token)
        {
            string host = "https://aip.baidubce.com/rest/2.0/ocr/v1/idcard?access_token=" + token;
            Encoding encoding = Encoding.Default;
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(host);
            request.Method = "post";
            request.KeepAlive = true;
            // 图片的base64编码
            string base64 = GetFileBase64(filePath);
            String str = "id_card_side=" + "front" + "&image=" + HttpUtility.UrlEncode(base64);
            byte[] buffer = encoding.GetBytes(str);
            request.ContentLength = buffer.Length;
            request.GetRequestStream().Write(buffer, 0, buffer.Length);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.Default);
            string result = reader.ReadToEnd();
            return result;
        }


        public static string GetFileBase64(string fileName)
        {
            FileStream filestream = new FileStream(fileName, FileMode.Open);
            byte[] arr = new byte[filestream.Length];
            filestream.Read(arr, 0, (int)filestream.Length);
            string baser64 = Convert.ToBase64String(arr);
            filestream.Close();
            return baser64;
        }

        public static string GetNumber(string jsonString)
        {
            var jsonNode = JsonSerializer.Deserialize<JsonObject>(jsonString);
            return jsonNode["words_result"]["公民身份号码"]["words"].ToString();
        }

        public static string GetName(string jsonString)
        {
            var jsonNode = JsonSerializer.Deserialize<JsonObject>(jsonString);
            return jsonNode["words_result"]["姓名"]["words"].ToString();
        }

        public static string GetStatus(string jsonString)
        {
            var jsonNode = JsonSerializer.Deserialize<JsonObject>(jsonString);
            return jsonNode["image_status"].ToString();
        }

        public static bool IsOK(string jsonString)
        {
            return GetStatus(jsonString).Equals("normal");
        }




    }

    public class FileInformation
    {
        public FileInformation(string filePath)
        {
            fileName = filePath;
        }
        public string fileName;
        public string GetFileLast()
        {
            string[] fileParts = fileName.Split('.');
            string[] fileArguments = fileParts[fileParts.Length - 2].Split('_');
            return fileArguments[fileArguments.Length - 1];
        }
        public int GetFileTime()
        {
            Console.WriteLine(this.GetFileLast());

            return int.Parse(this.GetFileLast());
        }
        public static int CompareFileDate(FileInformation file1, FileInformation file2)
        {
            if (file1.GetFileTime() > file2.GetFileTime())
            {
                return -1;
            }
            else if (file1.GetFileTime() < file2.GetFileTime())
            {
                return 1;
            }
            return 0;
        }
        public string GetFileName()
        {
            string name = fileName.Substring(fileName.LastIndexOf("\\") + 1);
            return name;
        }
    }
}
