using System;
using System.Globalization;
using System.Linq;
using System.Reflection;
using WixSharp;
using WixSharp.CommonTasks;
using WixSharp.Forms;

namespace ann.SetupWixSharp
{
    public class Script
    {
        static public void Main(string[] args)
        {
            //            var company = ((AssemblyCompanyAttribute)System.Reflection.Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyCompanyAttribute)).ElementAt(0)).Company;
            //            var product = ((AssemblyProductAttribute)System.Reflection.Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyProductAttribute)).ElementAt(0)).Product;
            var company = "ООО Надежда";
            var product = "YRL feed uploader";
            var version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString();

            const string releasePath = @"..\xlsx.convert.yrl\bin\Release";

            //Copy(@"..\..\..\DihvaEnoceanConfiguration\bin\Release\", "Release");

            var binaries = new Feature("Binaries", "Product binaries");

            var project = new ManagedProject(
//            var project = new Project(
                $"{product} (Version {version})",
                new Dir($@"%ProgramFiles%\{company}\{product}",
                    new File(binaries, $@"{releasePath}\xlsx.convert.yrl.exe.config"),
                    new File(binaries, $@"{releasePath}\DocumentFormat.OpenXml.dll"),
                    new File(binaries, $@"{releasePath}\NLog.config"),
                    new File(binaries, $@"{releasePath}\NLog.dll"),
                    new File(binaries, $@"{releasePath}\xlsx.convert.yrl.exe",
                        new FileShortcut(product, "INSTALLDIR"),
                        new FileShortcut(product, @"%Desktop%")
                    ),
                    new ExeFileShortcut($"Uninstall {product}", "[System64Folder]msiexec.exe", "/x [ProductCode]")
                ),

                new Dir($@"%Personal%\{product} - Photos"),

                new Dir($@"%ProgramMenu%\{company}\{product}",
                        new ExeFileShortcut($"{product} - Photos", "[" + $@"%Personal%\{product} - Photos".ToDirID() + "]", ""),
                        new ExeFileShortcut($"Uninstall {product}", "[System64Folder]msiexec.exe", "/x [ProductCode]")
                        ),
                new Property("ALLUSERS", "1")
            );

            project.GUID = new Guid("1E49E7CB-610D-447D-BD1A-C9A12359A35B");
            project.SetNetFxPrerequisite("NETFRAMEWORK45 >= '#379893'", "Please install .Net 4.5.2 First");
            project.BackgroundImage = @"files\logo_background.bmp";
            project.BannerImage = @"files\banner.png";
            project.LicenceFile = @"files\license.rtf";

            //            project.InstallScope = InstallScope.perUser; set privileges
                        project.Language = CultureInfo.CurrentCulture.Name;
                        project.Codepage = "1252";
            project.Description = "Yandex Realty converter Installer.";
            project.UpgradeCode = new Guid("1E49E7CB-610D-447D-BD1A-C9A12359A35B");
            project.MajorUpgrade = new MajorUpgrade() { /*DowngradeErrorMessage = $"A newer version of {product} is already installed."*/ };
            project.MajorUpgrade.AllowDowngrades = true;

            // custom set of standard UI dialogs
            project.ManagedUI = new ManagedUI();

            project.ManagedUI.InstallDialogs.Add(Dialogs.Welcome)
                                            .Add(Dialogs.Licence)
                                            .Add(Dialogs.InstallDir)
                                            .Add(Dialogs.Progress)
                                            .Add(Dialogs.Exit);
            Compiler.BuildMsi(project);

            //System.IO.Directory.Delete("Release", true);
        }
    }
}