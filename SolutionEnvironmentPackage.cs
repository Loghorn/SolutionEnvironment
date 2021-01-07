using EnvDTE;
using EnvDTE80;
using Microsoft;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using Task = System.Threading.Tasks.Task;

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
	[ProvideAutoLoad(UIContextGuids80.SolutionHasSingleProject, PackageAutoLoadFlags.BackgroundLoad)]
	[ProvideAutoLoad(UIContextGuids80.SolutionHasMultipleProjects, PackageAutoLoadFlags.BackgroundLoad)]
	[ProvideAutoLoad(VSConstants.UICONTEXT.SolutionOpening_string, PackageAutoLoadFlags.BackgroundLoad)]
	[ProvideAutoLoad(VSConstants.UICONTEXT.FullSolutionLoading_string, PackageAutoLoadFlags.BackgroundLoad)]
	[ProvideAutoLoad(VSConstants.UICONTEXT.Debugging_string, PackageAutoLoadFlags.BackgroundLoad)]
	[ProvideAutoLoad(VSConstants.UICONTEXT.SolutionBuilding_string, PackageAutoLoadFlags.BackgroundLoad)]
	[ProvideAutoLoad(VSConstants.UICONTEXT.SolutionExists_string, PackageAutoLoadFlags.BackgroundLoad)]
	[ProvideAutoLoad(VSConstants.UICONTEXT.SynchronousSolutionOperation_string, PackageAutoLoadFlags.BackgroundLoad)]
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
			await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

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

			_dte.Events.BuildEvents.OnBuildBegin += OnBuildBegin;

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

		private readonly List<EnvVar> savedEnvironmentVariables = new List<EnvVar>();

		string ResolveFilename(string filename, List<EnvVar> internalVars)
		{
			// Resolve all environment variables.
			filename = Regex.Replace(filename, @"\$\((\w+)\)", (m) =>
			{
				string env = m.Groups[1].Value;
				foreach (var v in internalVars)
				{
					if (string.Compare(env, v.name, true) == 0)
						return v.value;
				}
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

		bool ParseSlnEnvFile(string filename, List<EnvVar> internalVars, string solutionConfigurationName, string solutionPlatform)
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

					line = line.Trim();
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
								string newFilename = ResolveFilename(line.Substring(spacePos + 1), internalVars);
								if (!ParseSlnEnvFile(newFilename, internalVars, solutionConfigurationName, solutionPlatform))
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

					string name = line.Substring(0, equalsPos).Trim();

					int colonPos = name.IndexOf(':');
					if (colonPos != -1)
					{
						var match = Regex.Match(name.Substring(0, colonPos).Trim(), @"([^|]+)?(?:\|(.*))?");

						string configName = match.Groups[1].Value;
						string platform = match.Groups[2].Value;

						if (configName.Length > 0 && configName != solutionConfigurationName)
							continue;
						if (platform.Length > 0 && platform != solutionPlatform)
							continue;
						name = name.Substring(colonPos + 1).Trim();
					}

					string value = ResolveFilename(line.Substring(equalsPos + 1).Trim(), internalVars);

					string buffer = Environment.GetEnvironmentVariable(name);

					EnvVar envVar;
					envVar.name = name;
					envVar.value = buffer;
					savedEnvironmentVariables.Add(envVar);

					Environment.SetEnvironmentVariable(name, value);
					buffer = Environment.GetEnvironmentVariable(name);
					if (buffer != value)
						_pane.OutputString(string.Format("FAILED SETTING {0} = {1}\n", name, value));
					else
						_pane.OutputString(string.Format("SET {0} = {1}\n", name, value));
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
				string buffer = Environment.GetEnvironmentVariable(envVar.name);
				if (buffer != envVar.value)
					_pane.OutputString(string.Format("FAILED SETTING {0} = {1}\n", envVar.name, envVar.value));
				else
					_pane.OutputString(string.Format("SET {0} = {1}\n", envVar.name, envVar.value));
			}
		}

		public IEnumerable<T> Traverse<T>(IEnumerable<object> items, Func<object, IEnumerable<object>> childSelector)
		{
			var stack = new Stack<object>(items);
			while (stack.Any())
			{
				var next = stack.Pop();
				if (next is T)
					yield return (T)next;
				foreach (var child in childSelector(next))
					stack.Push(child);
			}
		}

		void ReadSolutionEnvironment()
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			string fullPath = _dte.Solution.FullName;

			try
			{
				UndoEnvironmentChanges();
				savedEnvironmentVariables.Clear();

				string solutionConfigurationName = _dte.Solution.SolutionBuild.ActiveConfiguration.Name;
				string solutionPlatform = "";

#pragma warning disable VSTHRD010 // Invoke single-threaded types on Main thread
				var configmgr = Traverse<Project>(_dte.Solution.Projects.Cast<Project>(), (p) =>
				{
					switch (p)
					{
						case Project project when project.ProjectItems != null:
							return project.ProjectItems.Cast<ProjectItem>();
						case ProjectItem item when item.SubProject != null:
							return Enumerable.Repeat(item.SubProject, 1);
					}
					return Enumerable.Empty<object>();
				}).Where((p) => p.ConfigurationManager != null).Select((p) => p.ConfigurationManager).FirstOrDefault();
#pragma warning restore VSTHRD010 // Invoke single-threaded types on Main thread
				if (configmgr != null)
				{
					var config = configmgr.ActiveConfiguration;
					solutionPlatform = config.PlatformName;
				}

				List<EnvVar> internalVars = new List<EnvVar>
				{
					new EnvVar{ name = "SolutionDir", value = Path.GetDirectoryName(fullPath) },
					new EnvVar{ name = "SolutionName", value = Path.GetFileNameWithoutExtension(fullPath) },
					new EnvVar{ name = "SolutionDrive", value = Path.GetPathRoot(fullPath) },
					new EnvVar{ name = "SolutionConfiguration", value = solutionConfigurationName },
					new EnvVar{ name = "SolutionPlatform", value = solutionPlatform },
				};

				fullPath = Path.ChangeExtension(fullPath, ".slnenv");

				if (ParseSlnEnvFile(fullPath, internalVars, solutionConfigurationName, solutionPlatform))
					_pane.OutputString(string.Format("Solution Environment {0} loaded\n", fullPath));
			}
			catch (Exception e)
			{
				_pane.OutputString(string.Format("Solution Environment {0} error: {1}\n", fullPath, e.Message));
			}
		}
	}
}
