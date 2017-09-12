using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using NLog;

namespace xlsx.convert.yrl
{
    public class CommandObject
    {
        [System.ComponentModel.DescriptionAttribute("file")]
        public string ExcelFile { get; set; }
        public int? SkipRows { get; set; }
        public string AgentOrganization { get; set; }
        public string AgentUrl { get; set; }
        public string AgentEmail { get; set; }
        public string AgentLogo { get; set; }
        public string[] AgentPhones { get; set; }
    }

    class Program
    {
        private static Logger _logger = LogManager.GetCurrentClassLogger();

        [STAThread]
        static void Main(string[] args)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = false;
            ofd.Title = "Выберите excel файл";
            if (ofd.ShowDialog() != DialogResult.OK)
            {
                return;
            }

//            var command = Args.Configuration.Configure<CommandObject>().CreateAndBind(args);
//            if (command.ExcelFile == null)
//            {
//                return;
//            }

            var salesAgent = new SalesAgent(
                string.Empty,
                Settings.Default.AgentOrganization,
                Settings.Default.AgentPhones.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries), 
                SalesAgent.CategoryType.Agency,
                Settings.Default.AgentUrl,
                Settings.Default.AgentEmail,
                Settings.Default.AgentLogo);

            string imagesFolder = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            IEnumerable<string> images = GetImages(imagesFolder);
            var converter = new ConvertExcelToYrl(_logger, salesAgent, images, Settings.Default.ImagesUri);

            XDocument doc = converter.GetXml(
                ofd.FileName,
                Settings.Default.SkipRows);

            if (converter.Counters.Exceptions.Any())
            {
                _logger.Warn("Converting errors: {0}", converter.Counters.Exceptions.Count());
            }
            else
            {
                doc.Save("export.xml");
                var client = new FtpWebClient(Settings.Default.ftpUrl, Settings.Default.ftpUser, Settings.Default.ftpPassword);
                client.ChangeWorkingDirectory("httpdocs");
                // Upload xml
                string[] items = client.ListDirectory();
                bool isExist = items.Any(i => i.Equals("httpdocs/hope-realty-yandex.xml"));
                if (isExist)
                {
                    client.DownloadFile("hope-realty-yandex.xml", $"hope-realty-yandex-{DateTime.Now.ToString("yyyyMMdd-hhmm")}-backup.xml");
                }
                string result = client.UploadFile("export.xml", "hope-realty-yandex.xml");
                _logger.Info($"ftp upload result: {result}");

                // Upload images
                if (!items.Any(i => i.Equals($"httpdocs/{Settings.Default.ftpImages}")))
                {
                    client.MakeDirectory(Settings.Default.ftpImages);
                }

                client.ChangeWorkingDirectory(Settings.Default.ftpImages);
                items = client.ListDirectory();
                foreach (var image in images)
                {
                    string destPath = $"{Settings.Default.ftpImages}/{image}";
                    if (!items.Any(i => i.Equals(destPath)))
                    {
                        result = client.UploadFile(Path.Combine(imagesFolder, "YRL feed uploader - Photos", image), image);
                        _logger.Info($"\t{image}: {result}");
                    }
                    else
                    {
                        _logger.Info($"\t{image}: exist");
                    }
                }
            }
            
            _logger.Info("finished");
        }

        private static IEnumerable<string> GetImages(string sourceDir)
        {
            string photosFolder = Path.Combine(sourceDir, "YRL feed uploader - Photos");

            DirectoryInfo d = new DirectoryInfo(photosFolder);
            FileInfo[] photos = d.GetFiles("*.*");
            return photos.Select(file => file.Name).ToList();
        } 
    }
}
