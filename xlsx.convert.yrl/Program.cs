using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

        static void Main(string[] args)
        {
            var command = Args.Configuration.Configure<CommandObject>().CreateAndBind(args);
            if (command.ExcelFile == null)
            {
                return;
            }

            var salesAgent = new SalesAgent(
                string.Empty,
                command.AgentOrganization, 
                command.AgentPhones, 
                SalesAgent.CategoryType.Agency, 
                command.AgentUrl, 
                command.AgentEmail,
                command.AgentLogo);

            var converter = new ConvertExcelToYrl(_logger, salesAgent);
            XDocument doc = converter.GetXml(
                command.ExcelFile, 
                command.SkipRows.HasValue ? command.SkipRows.Value : 0);

            if (converter.Counters.Exceptions.Any())
            {
                _logger.Warn("Converting errors: {0}", converter.Counters.Exceptions.Count());
            }
            else
            {
                doc.Save("export.xml");
            }
            
            _logger.Info("finished");
        }
    }
}
