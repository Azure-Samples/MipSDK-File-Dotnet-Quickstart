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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.InformationProtection.File;
using Microsoft.InformationProtection.Exceptions;
using Microsoft.InformationProtection;

namespace MipSdkDotNetQuickstart
{
    /// <summary>
    /// Action class implements the various MIP functionality.
    /// For this sample, only profile, engine, and handler creation are defined. 
    /// The IFileHandler may be used to label a file and read a labeled file.
    /// </summary>
    public class Action
    {
        private AuthDelegateImplementation authDelegate;
        private ApplicationInfo appInfo;
        private IFileProfile profile;
        private IFileEngine engine;
        private MipContext mipContext;

        // Used to pass in options for labeling the file.
        public struct FileOptions
        {
            public string FileName;
            public string OutputName;
            public string LabelId;
            public DataState DataState;
            public AssignmentMethod AssignmentMethod;
            public ActionSource ActionSource;
            public bool IsAuditDiscoveryEnabled;
            public bool GenerateChangeAuditEvent;
        }

        /// <summary>
        /// Constructor for Action class. Pass in AppInfo to simplify passing settings to AuthDelegate.
        /// </summary>
        /// <param name="appInfo"></param>
        public Action(ApplicationInfo appInfo)
        {
            this.appInfo = appInfo;

            // Initialize AuthDelegateImplementation using AppInfo. 
            authDelegate = new AuthDelegateImplementation(this.appInfo);

            // Initialize SDK DLLs. If DLLs are missing or wrong type, this will throw an exception

            MIP.Initialize(MipComponent.File);

            // This method in AuthDelegateImplementation triggers auth against Graph so that we can get the user ID.
            var id = authDelegate.GetUserIdentity();

            // Create profile.
            profile = CreateFileProfile(appInfo, ref authDelegate);

            // Create engine providing Identity from authDelegate to assist with service discovery.
            engine = CreateFileEngine(id);
        }

        /// <summary>
        /// Null refs to engine and profile and release all MIP resources.
        /// </summary>
        ~Action()
        {
            engine = null;
            profile = null;
            mipContext = null; 
        }

        /// <summary>
        /// Creates an IFileProfile and returns.
        /// IFileProfile is the root of all MIP SDK File API operations. Typically only one should be created per app.
        /// </summary>
        /// <param name="appInfo"></param>
        /// <param name="authDelegate"></param>
        /// <returns></returns>
        private IFileProfile CreateFileProfile(ApplicationInfo appInfo, ref AuthDelegateImplementation authDelegate)
        {
            // Initialize MipContext
            mipContext = MIP.CreateMipContext(appInfo, "mip_data", LogLevel.Trace, null, null);

                // Initialize file profile settings to create/use local state.                
                var profileSettings = new FileProfileSettings(mipContext, 
                    CacheStorageType.OnDiskEncrypted, 
                    authDelegate, 
                    new ConsentDelegateImplementation());

                // Use MIP.LoadFileProfileAsync() providing settings to create IFileProfile. 
                // IFileProfile is the root of all SDK operations for a given application.
                var profile = Task.Run(async () => await MIP.LoadFileProfileAsync(profileSettings)).Result;
                return profile;
            
        }

        /// <summary>
        /// Creates a file engine, associating the engine with the specified identity. 
        /// File engines are generally created per-user in an application. 
        /// IFileEngine implements all operations for fetching labels and sensitivity types.
        /// IFileHandlers are added to engines to perform labeling operations.
        /// </summary>
        /// <param name="identity"></param>
        /// <returns></returns>
        private IFileEngine CreateFileEngine(Identity identity)
        {

            // If the profile hasn't been created, do that first. 
            if (profile == null)
            {
                profile = CreateFileProfile(appInfo, ref authDelegate);
            }

            // Create file settings object. Passing in empty string for the first parameter, engine ID, will cause the SDK to generate a GUID.
            // Locale settings are supported and should be provided based on the machine locale, particular for client applications.
            var engineSettings = new FileEngineSettings("", "", "en-US")
            {
                // Provide the identity for service discovery.
                Identity = identity
            };

            // Add the IFileEngine to the profile and return.
            var engine = Task.Run(async () => await profile.AddEngineAsync(engineSettings)).Result;
            return engine;
        }
    

        /// <summary>
        /// Method creates a file handler and returns to the caller. 
        /// IFileHandler implements all labeling and protection operations in the File API.        
        /// </summary>
        /// <param name="options">Struct provided to set various options for the handler.</param>
        /// <returns></returns>
        private IFileHandler CreateFileHandler(FileOptions options)
        {            
                // Create the handler using options from FileOptions. Assumes that the engine was previously created and stored in private engine object.
                // There's probably a better way to pass/store the engine, but this is a sample ;)
                var handler = Task.Run(async () => await engine.CreateFileHandlerAsync(options.FileName, options.FileName, options.IsAuditDiscoveryEnabled)).Result;
                return handler;           
        }


        /// <summary>
        /// List all labels from the engine and return in IEnumerable<Label>
        /// </summary>
        /// <returns></returns>
        public IEnumerable<Label> ListLabels()
        {  
                // Get labels from the engine and return.
                // For a user principal, these will be user specific.
                // For a service principal, these may be service specific or global.
                return engine.SensitivityLabels;          
        }

        /// <summary>
        /// Set the label on the given file. 
        /// Options for the labeling operation are provided in the FileOptions parameter.
        /// </summary>
        /// <param name="options">Details about file input, output, label to apply, etc.</param>
        /// <returns></returns>
        public bool SetLabel(FileOptions options)
        {

            // LabelingOptions allows us to set the metadata associated with the labeling operations.
            // Review the API Spec at https://aka.ms/mipsdkdocs for details
            LabelingOptions labelingOptions = new LabelingOptions()
            {
                AssignmentMethod = options.AssignmentMethod
            };

            var handler = CreateFileHandler(options);

            // Use the SetLabel method on the handler, providing label ID and LabelingOptions
            // The handler already references a file, so those details aren't needed.
            handler.SetLabel(engine.GetLabelById(options.LabelId), labelingOptions, new ProtectionSettings());

            // The change isn't committed to the file referenced by the handler until CommitAsync() is called.
            // Pass the desired output file name in to the CommitAsync() function.
            var result = Task.Run(async () => await handler.CommitAsync(options.OutputName)).Result;

            // If the commit was successful and GenerateChangeAuditEvents is true, call NotifyCommitSuccessful()
            if (result && options.GenerateChangeAuditEvent)
            {
                // Submits and audit event about the labeling action to Azure Information Protection Analytics 
                handler.NotifyCommitSuccessful(options.FileName);
            }

            return result;
        }
                
        /// <summary>
        /// Read the label from a file provided via FileOptions.
        /// </summary>
        /// <param name="options"></param>
        /// <returns></returns>
        public ContentLabel GetLabel(FileOptions options)
        {           
                var handler = CreateFileHandler(options);
                return handler.Label;         
        }        
    }
}
