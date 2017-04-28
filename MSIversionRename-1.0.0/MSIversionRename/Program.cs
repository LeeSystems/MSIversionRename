using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using WindowsInstaller;


namespace LeeSystems.MSIversionRename
{
    class Program
    {

        static void Main(string[] args)
        {

            try
            {
                ParceCommandLine(args);

                if (SwitchDisplayHelp) Console.Write(GetHelp());
                // test if file exists
                if (!File.Exists(msiFile))
                {
                    Console.WriteLine("The file {0} Does not exist", msiFile);
                    AbortProcess = true;
                }
                // test for MSI extention
                String extention = Path.GetExtension(msiFile);
                if (!extention.Equals(".msi", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("The file must have a extention of 'MSI'");
                    AbortProcess = true;
                }
                if (!AbortProcess)
                {
                    if ((SwitchListAllProperties) && (NonSwitchArgumentIndex > 0)) DisplayAllProperties(msiFile);
                    if ((SwitchProcess) && (NonSwitchArgumentIndex > 0)) ProcessMSIfile(msiFile);
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            #if (DEBUG)
            Console.WriteLine("Debug Mode... Hit 'Enter' to exit");
            Console.Read();
            #endif
        }

        static void ProcessMSIfile(String MsiFile)
        {
            try
            {
                string version = "x.x.x";
                string productName = "[ProductName]";

                // Read the MSI property
                productName = GetMsiProperty(MsiFile, "ProductName");
                version = GetMsiProperty(MsiFile, "ProductVersion");
                String file = string.Format("{0}{1}{2}.msi", productName, NameSeperator,version);
                String OutputFile = Path.Combine(msiFilePath, file);
                Boolean overwrite = false;
                if (SwitchOverwriteFIle || SwitchRenameFIle) overwrite = true;
                File.Copy(MsiFile, OutputFile, overwrite);
                if (SwitchRenameFIle)
                {
                    File.Delete(MsiFile);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        static string GetMsiProperty(string msiFile, string property)
        {
            string retVal = string.Empty;

            // Create an Installer instance  
            Type classType = Type.GetTypeFromProgID("WindowsInstaller.Installer");
            Object installerObj = Activator.CreateInstance(classType);
            WindowsInstaller.Installer installer = installerObj as WindowsInstaller.Installer;

            // Open the msi file for reading  
            // 0 - Read, 1 - Read/Write  
            WindowsInstaller.Database database = installer.OpenDatabase(msiFile, 0);

            // Fetch the requested property  
            string sql = String.Format("SELECT Value FROM Property WHERE Property='{0}'", property);
            View view = database.OpenView(sql);
            view.Execute(null);

            // Read in the fetched record  
            Record record = view.Fetch();
            if (record != null)
            {
                retVal = record.get_StringData(1);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(record);
            }
            view.Close();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(view);
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(database);

            return retVal;
        }


        static void DisplayAllProperties(String msiFile)
        {
            Type classType = Type.GetTypeFromProgID("WindowsInstaller.Installer");
            Object installerObj = Activator.CreateInstance(classType);
            WindowsInstaller.Installer installer = installerObj as WindowsInstaller.Installer;

            // Open the msi file for reading  
            // 0 - Read, 1 - Read/Write  
            WindowsInstaller.Database database = installer.OpenDatabase(msiFile, 0);

            string sql = String.Format("SELECT Property, Value FROM Property");
            WindowsInstaller.View view = database.OpenView(sql);
               
            view.Execute(null);
            Record record = view.Fetch();
            while (record != null)
            {
                Console.WriteLine("{0} = {1}", record.get_StringData(1), record.get_StringData(2));
                record = view.Fetch();

            }
        }

        static Boolean SwitchRenameFIle = false;
        static Boolean SwitchOverwriteFIle = false;
        static Boolean SwitchListAllProperties = false;
        static Boolean SwitchDisplayHelp = false;
        static Boolean SwitchProcess = false;
        static Boolean AbortProcess = false;
        static String msiFile = String.Empty;
        static String msiFilePath = String.Empty;
        static String NameSeperator = @"-V";
        static int NonSwitchArgumentIndex = 0;

        static void ParceCommandLine(string[] args)
        {
            if ( args.Length == 0 )
            {
                SwitchDisplayHelp = true;
                return;
            }
            for (int i = 0; i < args.Length; i++)
            {
                String arg = args[i];
                switch (arg)
                {
                    case "/h":
                    case "-h":
                    case "-?":
                    case "-help":
                        {
                            SwitchDisplayHelp = true;
                            break;
                        }
                    case "-d":
                    case "-displayallproperties":
                        {
                            SwitchListAllProperties = true;
                            break;
                        }
                    case "-R":
                    case "-r":
                    case "-rename":
                        {
                            SwitchRenameFIle = true;
                            break;
                        }
                    case "-o":
                    case "-overwrite":
                        {
                            SwitchOverwriteFIle = true;
                            break;
                        }
                    case "-s":
                    case "-separator":
                        {
                            i++;
                            if ( i < args.Length)
                            {
                                NameSeperator = args[i];
                            }
                            break;
                        }
                    default:
                        {
                            if ( arg.StartsWith("-") || arg.StartsWith("/"))
                            {
                                Console.WriteLine("Argument Switch Unknown '{0}' Ignored",arg);
                                SwitchProcess = false;
                                AbortProcess = true;
                                break;
                            }
                            NonSwitchArgumentIndex++;
                            switch (NonSwitchArgumentIndex)
                            {
                                case 1:
                                    {
                                        msiFile = arg;
                                        msiFilePath = Path.GetDirectoryName(msiFile);
                                        SwitchProcess = true;
                                        break;
                                    }
                                default:
                                    {
                                        Console.WriteLine("Argument Number {0}Ignored: {1}", NonSwitchArgumentIndex, arg);
                                        SwitchProcess = false;
                                        break;
                                    }
                            }

                            break;
                        }
                }
            }
        }

        private static String GetHelp()
        {
            String help = Environment.NewLine;
            help += "Rename MSI Utility" + Environment.NewLine;
            help += "------------------" + Environment.NewLine;
            help += "Lee Systems, LLC" + Environment.NewLine;
            help += "www.leesystems.tv" + Environment.NewLine;
            help += "------------------" + Environment.NewLine;
            help += "Rename MSI file based on Product Name and Version Number" + Environment.NewLine;
            help += "Name based on 'Product Name' and 'Version' contained in MSI" + Environment.NewLine;
            help += "------------------------------------------------------------------------" + Environment.NewLine;
            help += "Command Line:" + Environment.NewLine;
            help += "++>MSIversionRename [options] [MSI filename]" + Environment.NewLine;
            help += "------------------------------------------------------------------------" + Environment.NewLine;
            help += "Command Line Options:" + Environment.NewLine;
            help += "     -help -h -? /h            > Display this text" + Environment.NewLine;
            help += "     -displayallproperties -d  > Display All properties in MSI" + Environment.NewLine;
            help += "     -rename -r -R             > rename original file. By default create" + Environment.NewLine;
            help += "     -                         > a new file (copy) with the new name" + Environment.NewLine;
            help += "     -overwrite -o             > Overwrite output file" + Environment.NewLine;
            help += "     -separator -s [separator] > Separator between Product Name and Version" + Environment.NewLine;
            help += "     -                         > Default separator is '-V'" + Environment.NewLine;
            help += "------------------------------------------------------------------------" + Environment.NewLine;

            return help;
        }
    }

}
