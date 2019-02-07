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
using System.Text.RegularExpressions;

namespace MipSdkDotNetQuickstart
{
    class Program
    {
        private const string clientId = "aed76448-c9a0-4035-b77c-1f3f4e69a489";
        private const string appName = "Sdk Test App";


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
            string inputFilePath = @"D:\MIP\TestFiles\Public.docx";            
            string outputFilePath = @"D:\MIP\TestFiles\Public_modified.docx";

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
