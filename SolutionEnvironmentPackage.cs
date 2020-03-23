﻿using EnvDTE;
using EnvDTE80;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft;
using System.Collections.Generic;
using Microsoft.Win32;
using System.IO;
using Microsoft.VisualStudio;
using System.Text.RegularExpressions;

namespace SolutionEnvironment
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(SolutionEnvironmentPackage.PackageGuidString)]
    //[ProvideAutoLoad(UIContextGuids80.SolutionHasSingleProject, PackageAutoLoadFlags.BackgroundLoad)]
    //[ProvideAutoLoad(UIContextGuids80.SolutionHasMultipleProjects, PackageAutoLoadFlags.BackgroundLoad)]
    [ProvideAutoLoad(VSConstants.UICONTEXT.SolutionOpening_string, PackageAutoLoadFlags.BackgroundLoad)]
    public sealed class SolutionEnvironmentPackage : AsyncPackage
    {
        private DTE2 _dte;
        private IVsOutputWindowPane _pane;

        /// <summary>
        /// SolutionEnvironmentPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "963e6582-9704-4ac0-83c1-14d82bde9d6a";

        #region Package Members

        /// <summary>
        /// Initialization of the package; this method is called right after the package is sited, so this is the place
        /// where you can put all the initialization code that rely on services provided by VisualStudio.
        /// </summary>
        /// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
        /// <param name="progress">A provider for progress updates.</param>
        /// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            // When initialized asynchronously, the current thread may be a background thread at this point.
            // Do any initialization that requires the UI thread after switching to the UI thread.
            await this.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            _dte = await GetServiceAsync(typeof(DTE)) as DTE2;
            Assumes.Present(_dte);

            _pane = CreatePane(new Guid(), "Solution Environment", true, true);
            _pane.OutputString("Solution Environment loaded\n");

            var solService = await GetServiceAsync(typeof(SVsSolution)) as IVsSolution;
            Assumes.Present(solService);
            ErrorHandler.ThrowOnFailure(solService.GetProperty((int)__VSPROPID.VSPROPID_IsSolutionOpen, out object value));
            if (value is bool isSolOpen && isSolOpen)
            {
                ReadSolutionEnvironment();
            }

            _dte.Events.DebuggerEvents.OnEnterRunMode += OnEnterRunMode;

            _dte.Events.BuildEvents.OnBuildBegin += OnBuildBegin; ;

            _dte.Events.SolutionEvents.Opened += SolutionOpened;
        }

        private void OnBuildBegin(vsBuildScope Scope, vsBuildAction Action)
        {
            JoinableTaskFactory.Run(async () =>
            {
                await JoinableTaskFactory.SwitchToMainThreadAsync();
                ReadSolutionEnvironment();
            });
        }

        private void SolutionOpened()
        {
            JoinableTaskFactory.Run(async () =>
            {
                await JoinableTaskFactory.SwitchToMainThreadAsync();
                ReadSolutionEnvironment();
            });
        }

        private void OnEnterRunMode(dbgEventReason Reason)
        {
            if (Reason == dbgEventReason.dbgEventReasonLaunchProgram)
            {
                JoinableTaskFactory.Run(async () =>
                {
                    await JoinableTaskFactory.SwitchToMainThreadAsync();
                    ReadSolutionEnvironment();
                });
            }
        }

        #endregion

        IVsOutputWindowPane CreatePane(Guid paneGuid, string title, bool visible, bool clearWithSolution)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            IVsOutputWindow output = (IVsOutputWindow)GetService(typeof(SVsOutputWindow));
            Assumes.Present(output);

            IVsOutputWindowPane pane;

            // Create a new pane.
            output.CreatePane(
                ref paneGuid,
                title,
                Convert.ToInt32(visible),
                Convert.ToInt32(clearWithSolution));

            // Retrieve the new pane.
            output.GetPane(ref paneGuid, out pane);

            return pane;
        }

        private struct EnvVar
        {
            public string name;
            public string value;
        };

        private List<EnvVar> savedEnvironmentVariables = new List<EnvVar>();

        string ResolveFilename(string filename, string solutionDir, string solutionName)
        {
            // Resolve all environment variables.
            filename = Regex.Replace(filename, @"\$\((\w+)\)", (m) =>
            {
                string env = m.Groups[1].Value;
                // See if we can resolve it.  If not, then exit.
                if (string.Compare(env, "solutiondir", true) == 0)
                    return solutionDir;
                else if (string.Compare(env, "solutionname", true) == 0)
                    return solutionName;
                else
                    return Environment.GetEnvironmentVariable(env);
            });

            // Resolve all registry entries.
            filename = Regex.Replace(filename, @"%\(([^\)]+)\)", (m) =>
            {
                string reg = m.Groups[1].Value;

                string root = reg.Substring(0, 4);
                reg = reg.Substring(5);

                // Strip the key.
                int slashPos = reg.LastIndexOf('\\');
                if (slashPos != -1)
                {
                    string key = reg.Substring(slashPos + 1);
                    reg = reg.Substring(0, slashPos);

                    RegistryKey parentKey = null;
                    if (root == "HKLM")
                        parentKey = Registry.LocalMachine;
                    else if (root == "HKCU")
                        parentKey = Registry.CurrentUser;

                    var regKey = parentKey?.OpenSubKey(reg);
                    return regKey?.GetValue(key, "") as string ?? "";
                }

                return "";
            });

            return filename;
        }

        bool ParseSlnEnvFile(string filename, string solutionDir, string solutionName, string solutionConfigurationName)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!File.Exists(filename))
                return false;

            int lineNumber = 0;

            using (var reader = new StreamReader(filename))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    lineNumber++;

                    line.Trim();
                    if (line.Length == 0)
                        continue;

                    // Comment
                    if (line.StartsWith("--"))
                        continue;

                    int equalsPos = line.IndexOf('=');
                    if (equalsPos == -1)
                    {
                        int spacePos = line.IndexOf(' ');
                        if (spacePos != -1)
                        {
                            string command = line.Substring(0, spacePos).ToLower();
                            if (command == "include" || command == "forceinclude")
                            {
                                string newFilename = ResolveFilename(line.Substring(spacePos + 1), solutionDir, solutionName);
                                if (!ParseSlnEnvFile(newFilename, solutionDir, solutionName, solutionConfigurationName))
                                {
                                    if (command == "forceinclude")
                                    {
                                        throw new Exception($"Line #{lineNumber} of file [{filename}] is invalid.  include [{newFilename}.slnenv] not found.");
                                    }
                                }
                                continue;
                            }
                            else
                            {
                                throw new Exception($"Line #{lineNumber} of file [{filename}] is invalid.  include Filename or name=value expected.");
                            }
                        }
                        else
                        {
                            throw new Exception($"Line #{lineNumber} of file [{filename}] is invalid.  include Filename or name=value expected.");
                        }
                    }

                    string name = line.Substring(0, equalsPos);
                    name.Trim();

                    int colonPos = name.IndexOf(':');
                    if (colonPos != -1)
                    {
                        string configName = name.Substring(0, colonPos);
                        configName.Trim();
                        if (configName != solutionConfigurationName)
                            continue;
                        name = name.Substring(colonPos + 1);
                        name.Trim();
                    }

                    string value = ResolveFilename(line.Substring(equalsPos + 1).Trim(), solutionDir, solutionName);

                    string buffer = Environment.GetEnvironmentVariable(name);

                    EnvVar envVar;
                    envVar.name = name;
                    envVar.value = buffer;
                    savedEnvironmentVariables.Add(envVar);

                    Environment.SetEnvironmentVariable(name, value);
                    _pane.OutputString(string.Format("{0} = {1}\n", name, value));
                }

                return true;
            }
        }

        void UndoEnvironmentChanges()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            foreach (var envVar in savedEnvironmentVariables)
            {
                Environment.SetEnvironmentVariable(envVar.name, envVar.value);
                _pane.OutputString(string.Format("{0} = {1}\n", envVar.name, envVar.value));
            }
        }

        void ReadSolutionEnvironment()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            UndoEnvironmentChanges();
            savedEnvironmentVariables.Clear();

            string fullPath = _dte.Solution.FullName;

            string solutionDir = Path.GetDirectoryName(fullPath);

            string solutionName = Path.GetFileNameWithoutExtension(fullPath);

            string solutionConfigurationName = _dte.Solution.SolutionBuild.ActiveConfiguration.Name;

            fullPath = Path.ChangeExtension(fullPath, ".slnenv");

            if (ParseSlnEnvFile(fullPath, solutionDir, solutionName, solutionConfigurationName))
                _pane.OutputString(string.Format("Solution Environment {0} loaded\n", fullPath));
        }
    }
}