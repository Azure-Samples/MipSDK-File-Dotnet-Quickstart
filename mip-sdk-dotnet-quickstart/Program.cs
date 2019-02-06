using System;
using System.Threading.Tasks;
using Microsoft.InformationProtection;
using Microsoft.InformationProtection.File;
using Microsoft.InformationProtection.Protection;
using Microsoft.InformationProtection.Policy;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace MipSdkDotNetQuickstart
{
    class Program
    {
        private const string clientId = "YOUR CLIENT ID";
        private const string appName = "YOUR APP NAME";


        static void Main(string[] args)
        { 
            // Create ApplicationInfo, setting the clientID from Azure AD App Registration as the ApplicationId
            ApplicationInfo appInfo = new ApplicationInfo()
            {
                ApplicationId = clientId,
                ApplicationName = appName,
                ApplicationVersion = "1.0.0"
            };

            // Initialize Action class, passing in AppInfo.
            Action action = new Action(appInfo);

            // List all labels available to the engine created in Action
            IEnumerable<Label> labels = action.ListLabels();


            // Enumerate parent and child labels and print name/ID. 
            foreach (var label in labels)
            {
                Console.WriteLine(string.Format("{0} - {1}", label.Name, label.Id));

                if (label.Children.Count > 0)
                {
                    foreach (Label child in label.Children)
                    {
                        Console.WriteLine(string.Format("\t{0} - {1}", child.Name, child.Id));
                    }
                }
            }

            // Prompt user to enter a label ID from above
            Console.Write("Enter a label identifier from above: ");
            var labelId = Console.ReadLine();
                        

            // Set paths and label ID
            string inputFilePath = @"D:\MIP\TestFiles\demo1.docx";            
            string outputFilePath = @"D:\MIP\TestFiles\Test1_modified.docx";

            // Set file options from FileOptions struct. Used to set various parameters for FileHandler
            Action.FileOptions options = new Action.FileOptions
            {
                FileName = inputFilePath,
                OutputName = outputFilePath,
                ActionSource = ActionSource.Manual,
                AssignmentMethod = AssignmentMethod.Standard,
                ContentState = ContentState.Rest,
                GenerateChangeAuditEvent = true,
                IsAuditDiscoveryEnabled = true,
                LabelId = labelId
            };
            
            //Set the label on the file handler object
            Console.WriteLine(string.Format("Set label ID {0} on {1}", labelId, inputFilePath));

            // Set label, commit change to outputfile, and send audit event if enabled.
            var result = action.SetLabel(options);
            
            Console.WriteLine(string.Format("Committed label ID {0} to {1}", labelId, outputFilePath));

            // Create a new handler to read the labeled file metadata.           
            Console.WriteLine(string.Format("Getting the label committed to file: {0}", outputFilePath));

            // Update options to read the previously generated file output.
            options.FileName = options.OutputName;

            // Read label from the previously labeled file.
            var contentLabel = action.GetLabel(options);
                                                    
            // Display the label with protection information.
            Console.WriteLine(string.Format("File Label: {0} \r\nIsProtected: {1}", contentLabel.Label.Name, contentLabel.IsProtectionAppliedFromLabel.ToString()));
            Console.WriteLine("Press a key to quit.");
            Console.ReadKey();
        }
    }
}
