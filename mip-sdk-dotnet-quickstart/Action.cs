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

        // Used to pass in options for labeling the file.
        public struct FileOptions
        {
            public string FileName;
            public string OutputName;
            public string LabelId;
            public ContentState ContentState;
            public AssignmentMethod AssignmentMethod;
            public ActionSource ActionSource;
            public bool IsAuditDiscoveryEnabled;
            public bool GenerateChangeAuditEvent;
        }

        // Accept ApplicationInfo in contructor parameter. AppInfo used to initialize the AuthDelegate.
        public Action(ApplicationInfo appInfo)
        {
            this.appInfo = appInfo;

            // Initialize AuthDelegateImplementation using AppInfo. 
            authDelegate = new AuthDelegateImplementation(this.appInfo);

            // Initialize SDK DLLs. If DLLs are missing or wrong type, this will throw an exception
            try
            {
                MIP.Initialize(MipComponent.File);
                
                // This method in AuthDelegateImplementation triggers auth against Graph so that we can get the user ID
                var id = authDelegate.GetUserIdentity();
                profile = CreateFileProfile(appInfo, ref authDelegate);
                engine = CreateFileEngine(id);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private IFileProfile CreateFileProfile(ApplicationInfo appInfo, ref AuthDelegateImplementation authDelegate)
        {
            try
            {
                // Initialize file profile settings to create/use local state.
                var profileSettings = new FileProfileSettings("mip_data", false, authDelegate, new ConsentDelegateImplementation(), appInfo, LogLevel.Trace);
                var profile = Task.Run(async () => await MIP.LoadFileProfileAsync(profileSettings)).Result;
                return profile;
            }

            catch (Exception ex)
            {
                throw ex;
            }
        }

        private IFileEngine CreateFileEngine(Identity identity)
        {
            try
            {
                if(profile == null)
                {
                    profile = CreateFileProfile(appInfo, ref authDelegate);
                }

                var engineSettings = new FileEngineSettings("", "", "en-US")
                {
                    Identity = identity
                };

                var engine = Task.Run(async () => await profile.AddEngineAsync(engineSettings)).Result;
                return engine;
            }

            catch (Exception ex)
            {
                throw ex;
            }
        }

        private IFileHandler CreateFileHandler(FileOptions options)
        {
            try
            {
                if(engine == null)
                {
                    //CreateFileEngine()
                }

                var handler = Task.Run(async () => await engine.CreateFileHandlerAsync(options.FileName, options.FileName, options.ContentState, options.IsAuditDiscoveryEnabled)).Result;
                return handler;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public IEnumerable<Label> ListLabels()
        {
            try
            {
                if (engine == null)
                {
                    //CreateFileEngine(id);
                }

                return engine.SensitivityLabels;
            }

            catch (Exception ex)
            {
                throw ex;
            }
        }

        public bool SetLabel(FileOptions options)
        {
            try
            {
                LabelingOptions labelingOptions = new LabelingOptions()
                {
                    ActionSource = options.ActionSource,
                    AssignmentMethod = options.AssignmentMethod
                };

                var handler = CreateFileHandler(options);

                handler.SetLabel(options.LabelId, labelingOptions);

                var result = Task.Run(async() => await handler.CommitAsync(options.OutputName)).Result;

                if(result && options.GenerateChangeAuditEvent)
                {
                    handler.NotifyCommitSuccessful(options.FileName);
                }

                return result;
            }

            catch(Exception ex)
            {
                return false;
            }
        }

        public ContentLabel GetLabel(FileOptions options)
        {
            try
            {
                var handler = CreateFileHandler(options);
                return handler.Label;
            }
            catch (Exception ex)
            {

                throw ex;
            }
        }
        
    }
}
