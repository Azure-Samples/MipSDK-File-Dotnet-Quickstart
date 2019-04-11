/*
*
* Copyright (c) Microsoft Corporation.
* All rights reserved.
*
* This code is licensed under the MIT License.
*
* Permission is hereby granted, free of charge, to any person obtaining a copy
* of this software and associated documentation files(the "Software"), to deal
* in the Software without restriction, including without limitation the rights
* to use, copy, modify, merge, publish, distribute, sublicense, and / or sell
* copies of the Software, and to permit persons to whom the Software is
* furnished to do so, subject to the following conditions :
*
* The above copyright notice and this permission notice shall be included in
* all copies or substantial portions of the Software.
*
* THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
* IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
* FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.IN NO EVENT SHALL THE
* AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
* LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
* OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
* THE SOFTWARE.
*
*/

using System;
using System.Threading.Tasks;
using Microsoft.InformationProtection;
using Microsoft.InformationProtection.File;
using Microsoft.InformationProtection.Protection;
using Microsoft.InformationProtection.Policy;
using System.Collections.Generic;
using System.Configuration;

namespace MipSdkDotNetQuickstart
{
    class Program
    {
        private static readonly string clientId = ConfigurationManager.AppSettings["ida:ClientId"];
        private static readonly string appName = ConfigurationManager.AppSettings["app:Name"];
        private static readonly string appVersion = ConfigurationManager.AppSettings["app:Version"];

        static void Main(string[] args)
        {
            try
            {
                // Create ApplicationInfo, setting the clientID from Azure AD App Registration as the ApplicationId
                // If any of these values are not set API throws BadInputException.
                ApplicationInfo appInfo = new ApplicationInfo()
                {
                    // ApplicationId should ideally be set to the same ClientId found in the Azure AD App Registration.
                    // This ensures that the clientID in AAD matches the AppId reported in AIP Analytics.
                    ApplicationId = clientId,
                    ApplicationName = appName,
                    ApplicationVersion = appVersion
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

                // Prompt for path inputs
                Console.Write("Enter an input file path: ");
                string inputFilePath = Console.ReadLine();

                Console.Write("Enter an output file path: ");
                string outputFilePath = Console.ReadLine();

                // Set file options from FileOptions struct. Used to set various parameters for FileHandler
                Action.FileOptions options = new Action.FileOptions
                {
                    FileName = inputFilePath,
                    OutputName = outputFilePath,
                    ActionSource = ActionSource.Manual,
                    AssignmentMethod = AssignmentMethod.Standard,
                    DataState = DataState.Rest,
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
            catch (Exception ex)
            {
                Console.WriteLine("****** Error!");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);                
            }

        }
    }
}
