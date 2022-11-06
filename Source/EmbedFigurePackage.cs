/*
 * EmbedFigure - Visual Studio extension for embedding math figures into source code
 * Copyright(C) 2022 Tamas Kezdi
 *
 * This program is free software : you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program. If not, see <https://www.gnu.org/licenses/>.
 */

#if TRACE
//#define TRACE_FUNCTIONS
#endif

using MVS   = Microsoft.VisualStudio;
using MVSS  = Microsoft.VisualStudio.Shell;
using MVSSI = Microsoft.VisualStudio.Shell.Interop;
using MVSSS = Microsoft.VisualStudio.Shell.Settings;
using MVSSe = Microsoft.VisualStudio.Settings;
using MVST  = Microsoft.VisualStudio.Threading;
using S     = System;
using SCG   = System.Collections.Generic;
using SG    = System.Globalization;
using SCM   = System.ComponentModel;
using SRIS  = System.Runtime.InteropServices;
using ST    = System.Threading;
using STT   = System.Threading.Tasks;

namespace EmbedFigure
{
	[SRIS.Guid(c_OptionsGuidString)]
	internal class EmbedFigureOptions
	{
		#region Constants

		private const int c_DefaultUpdateDelay = 1500;
		private const char c_DefaultPrefixChar = '@';
		/// <summary>
		/// Registry key name where the setting are stored.
		/// </summary>
		/// <remarks>
		/// This registry key is stored in the private registry of the Visual Studio.
		/// The private registry file is here: C:\Users\<user_name>\AppData\Local\Microsoft\VisualStudio\<version>\privateregistry.bin
		/// It can be loaded into the registry editor with "Load Hive..."
		/// </remarks>
		private const string c_CollectionName = "EmbedFigure_1_57EFFFD0-11F4-4987-8E0B-58A55024D9CA";
		public const string c_OptionsGuidString = "244FF672-5B6D-4657-8B99-BC288D3A1DC9";

		internal const int c_Version_Major = 0;
		internal const int c_Version_Minor = 1;
		internal const int c_Version_Build = 0;
		internal const int c_Version_Revision = 0;

		#endregion

		#region Static variables

		private static readonly char[] s_ValidPrefixChars = { '@', '#' };

		private static MVST.AsyncLazy<EmbedFigureOptions> s_Options = new MVST.AsyncLazy<EmbedFigureOptions>(CreateOptionsAsync, MVSS.ThreadHelper.JoinableTaskFactory);
		private static MVST.AsyncLazy<MVSSS.ShellSettingsManager> s_SettingsManager = new MVST.AsyncLazy<MVSSS.ShellSettingsManager>(GetSettingsManagerAsync, MVSS.ThreadHelper.JoinableTaskFactory);

		#endregion

		#region Static functions

		private static async STT.Task<EmbedFigureOptions> CreateOptionsAsync()
		{
			var options = new EmbedFigureOptions();
			await options.LoadOptionsAsync();
			return options;
		}

		private static async STT.Task<MVSSS.ShellSettingsManager> GetSettingsManagerAsync()
		{
			// Switch to main thread, because IVsSettingsManager can only be used on the main thread
			await MVSS.ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
			var vs_settings_manager = await MVSS.AsyncServiceProvider.GlobalProvider.GetServiceAsync(typeof(MVSSI.SVsSettingsManager)) as MVSSI.IVsSettingsManager;
			if (null == vs_settings_manager)
			{
				return null;
			}
			return new MVSSS.ShellSettingsManager(vs_settings_manager);
		}

		private static bool IsValidPrefixChar(char prefix_char)
		{
			foreach (char current_prefix_char in s_ValidPrefixChars)
			{
				if (current_prefix_char == prefix_char)
				{
					return true;
				}
			}
			return false;
		}
		#endregion

		#region Static properties

		internal static EmbedFigureOptions Instance
		{
			get => MVSS.ThreadHelper.JoinableTaskFactory.Run(s_Options.GetValueAsync);
		}

		#endregion

		#region Member variables

		// Current values of settings
		private int m_UpdateDelay = c_DefaultUpdateDelay;
		private char m_PrefixChar = c_DefaultPrefixChar;

		private int m_LastSavedUpdateDelay = c_DefaultUpdateDelay;
		private char m_LastSavedPrefixChar = c_DefaultPrefixChar;

		internal int m_PrevUpdateDelay = c_DefaultUpdateDelay;
		internal char m_PrevPrefixChar = c_DefaultPrefixChar;

		#endregion

		#region Events

		internal event S.EventHandler<S.EventArgs> OptionsChanged;

		#endregion

		#region Member properties

		[SCM.Category("General")]
		[SCM.Description("Update figures after this much time has elapsed since the last change in ms.")]
		[SCM.DisplayName("Update delay")]
		public int UpdateDelay
		{
			get => m_UpdateDelay;
			set => m_UpdateDelay = value;
		}

		[SCM.Category("General")]
		[SCM.Description("The character before EmbedFigure.")]
		[SCM.DisplayName("Prefix character")]
		[SCM.TypeConverter(typeof(PrefixCharTypeConverter))]
		public char PrefixChar
		{
			get => m_PrefixChar;
			set => m_PrefixChar = value;
		}

		#endregion

		#region Member functions

		private async STT.Task LoadOptionsAsync()
		{
#if TRACE_FUNCTIONS
			Trace.EnterFunction();
#endif
			MVSSS.ShellSettingsManager settings_manager = await s_SettingsManager.GetValueAsync();
			if (null == settings_manager)
			{
				ResetSettings();
#if TRACE_FUNCTIONS
				Trace.LeaveFunction();
#endif
				return;
			}

			MVSSe.SettingsStore settings_store = settings_manager.GetReadOnlySettingsStore(MVSSe.SettingsScope.UserSettings);
			if (!settings_store.CollectionExists(c_CollectionName))
			{
				ResetSettings();
#if TRACE_FUNCTIONS
				Trace.LeaveFunction();
#endif
				return;
			}

			try
			{
				m_UpdateDelay = m_PrevUpdateDelay = m_LastSavedUpdateDelay = settings_store.GetInt32(c_CollectionName, "UpdateDelay");
			}
			catch
			{
				m_UpdateDelay = m_PrevUpdateDelay = m_LastSavedUpdateDelay = c_DefaultUpdateDelay;
			}

			try
			{
				string prefix_char = settings_store.GetString(c_CollectionName, "PrefixChar");
				if (1 == prefix_char.Length && IsValidPrefixChar(prefix_char[0]))
				{
					m_PrefixChar = m_PrevPrefixChar = m_LastSavedPrefixChar = prefix_char[0];
				}
				else
				{
					m_PrefixChar = m_PrevPrefixChar = m_LastSavedPrefixChar = c_DefaultPrefixChar;
				}
			}
			catch
			{
				m_PrefixChar = m_PrevPrefixChar = m_LastSavedPrefixChar = c_DefaultPrefixChar;
			}
#if TRACE_FUNCTIONS
			Trace.LeaveFunction();
#endif
		}

		private async STT.Task SaveOptionsAsync()
		{
#if TRACE_FUNCTIONS
			Trace.EnterFunction();
#endif
			MVSSS.ShellSettingsManager settings_manager = await s_SettingsManager.GetValueAsync();
			MVSSe.WritableSettingsStore writable_settings_store = null;
			if (null != settings_manager)
			{
				try
				{
					writable_settings_store = settings_manager.GetWritableSettingsStore(MVSSe.SettingsScope.UserSettings);
				}
				catch
				{
				}

				if (null != writable_settings_store && !writable_settings_store.CollectionExists(c_CollectionName))
				{
					writable_settings_store.CreateCollection(c_CollectionName);
				}
			}

			bool options_changed = false;

			m_PrevUpdateDelay = m_LastSavedUpdateDelay;
			if (m_UpdateDelay != m_LastSavedUpdateDelay)
			{
				options_changed = true;
				m_LastSavedUpdateDelay = m_UpdateDelay;
				if (null != writable_settings_store)
				{
					try
					{
						writable_settings_store.SetInt32(c_CollectionName, "UpdateDelay", m_LastSavedUpdateDelay);
					}
					catch
					{
					}
				}
			}

			m_PrevPrefixChar = m_LastSavedPrefixChar;
			if (m_PrefixChar != m_LastSavedPrefixChar)
			{
				options_changed = true;
				m_LastSavedPrefixChar = m_PrefixChar;
				if (null != writable_settings_store)
				{
					try
					{
						writable_settings_store.SetString(c_CollectionName, "PrefixChar", m_LastSavedPrefixChar.ToString());
					}
					catch
					{
					}
				}
			}

			if (options_changed)
			{
				OptionsChanged?.Invoke(this, null);
			}
#if TRACE_FUNCTIONS
			Trace.LeaveFunction();
#endif
		}

		internal void LoadSettingsFromStorage()
		{
			MVSS.ThreadHelper.JoinableTaskFactory.Run(LoadOptionsAsync);
		}

		internal void LoadSettingsFromXml(MVSSI.IVsSettingsReader reader)
		{
#if TRACE_FUNCTIONS
			Trace.EnterFunction();
#endif
			MVSS.ThreadHelper.JoinableTaskFactory.Run(async delegate
			{
				await MVSS.ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

#if TRACE_FUNCTIONS
				Trace.Message("Switched to Main LoadSettingsFromXml");
#endif
				bool options_changed = false;

				if (MVS.VSConstants.S_OK == reader.ReadCategoryVersion(out int version_major, out int version_minor, out int version_build, out int version_revision))
				{
					if (c_Version_Major == version_major && c_Version_Minor == version_minor && c_Version_Build == version_build && c_Version_Revision == version_revision)
					{
						if (MVS.VSConstants.S_OK == reader.ReadSettingLong("UpdateDelay", out int update_delay))
						{
							if (update_delay != m_UpdateDelay)
							{
								m_UpdateDelay = update_delay;
								options_changed = true;
							}
						}
						if (MVS.VSConstants.S_OK == reader.ReadSettingString("PrefixChar", out string prefix_char))
						{
							if (null != prefix_char && 1 == prefix_char.Length && IsValidPrefixChar(prefix_char[0]) && prefix_char[0] != m_PrefixChar)
							{
								m_PrefixChar = prefix_char[0];
								options_changed = true;
							}
						}
					}
				}

				if (options_changed)
				{
					OptionsChanged?.Invoke(this, null);
				}

#if TRACE_FUNCTIONS
				Trace.Message("Switch from Main LoadSettingsFromXml");
#endif
			});
#if TRACE_FUNCTIONS
			Trace.LeaveFunction();
#endif
		}

		internal void SaveSettingsToStorage()
		{
			MVSS.ThreadHelper.JoinableTaskFactory.Run(SaveOptionsAsync);
		}

		internal void SaveSettingsToXml(MVSSI.IVsSettingsWriter writer)
		{
#if TRACE_FUNCTIONS
			Trace.EnterFunction();
#endif
			MVSS.ThreadHelper.JoinableTaskFactory.Run(async delegate
			{
				await MVSS.ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

#if TRACE_FUNCTIONS
				Trace.Message("Switched to Main SaveSettingsToXml");
#endif
				writer.WriteCategoryVersion(c_Version_Major, c_Version_Minor, c_Version_Build, c_Version_Revision);
				writer.WriteSettingLong("UpdateDelay", m_UpdateDelay);
				writer.WriteSettingString("PrefixChar", m_PrefixChar.ToString());
#if TRACE_FUNCTIONS
				Trace.Message("Switch from Main SaveSettingsToXml");
#endif
			});
#if TRACE_FUNCTIONS
			Trace.LeaveFunction();
#endif
		}

		internal void ResetSettings()
		{
			m_LastSavedUpdateDelay = m_UpdateDelay = c_DefaultUpdateDelay;
			m_LastSavedPrefixChar = m_PrefixChar = c_DefaultPrefixChar;
		}

		#endregion
	}

	/// <summary>
	/// Specifying this <see cref="SCM.TypeConverter"/> for the <see cref="EmbedFigureOptionPageGrid.PrefixChar"/> creates a drop-down list to select from the allowed characters.
	/// </summary>
	internal class PrefixCharTypeConverter : SCM.TypeConverter
	{
		/// <summary>
		/// Indicates that prefix char can be converted from string.
		/// </summary>
		/// <remarks>This function is called by the framework on Main Thread</remarks>
		public override bool CanConvertFrom(SCM.ITypeDescriptorContext context, S.Type source_type)
		{
			if (typeof(string) == source_type)
			{
				return true;
			}
			return base.CanConvertFrom(context, source_type);
		}

		/// <summary>
		/// Indicates that prefix char can be converted to string.
		/// </summary>
		/// <remarks>This function is called by the framework on Main Thread</remarks>
		public override bool CanConvertTo(SCM.ITypeDescriptorContext context, S.Type destination_type)
		{
			if (typeof(string) == destination_type)
			{
				return true;
			}
			return base.CanConvertTo(context, destination_type);
		}

		/// <summary>
		/// Implements string to char conversion
		/// </summary>
		/// <remarks>This function is called by the framework on Main Thread</remarks>
		public override object ConvertFrom(SCM.ITypeDescriptorContext context, SG.CultureInfo culture, object value)
		{
			if (value is string)
			{
				// Compare strings by content
				if (value.Equals("#"))
				{
					return '#';
				}
				return '@';
			}
			return base.ConvertFrom(context, culture, value);
		}

		/// <summary>
		/// Implements char to string conversion
		/// </summary>
		/// <remarks>This function is called by the framework on Main Thread</remarks>
		public override object ConvertTo(SCM.ITypeDescriptorContext context, SG.CultureInfo culture, object value, S.Type destination_type)
		{
			if (destination_type == typeof(string))
			{
				if ('#' == (char)value)
				{
					return "#";
				}
				return "@";
			}
			return base.ConvertTo(context, culture, value, destination_type);
		}

		/// <summary>
		/// Specifies the content of the drop-down list.
		/// </summary>
		/// <remarks>This function is called by the framework on Main Thread</remarks>
		public override StandardValuesCollection GetStandardValues(SCM.ITypeDescriptorContext context)
		{
			var list = new SCG.List<char>();
			list.Add('@');
			list.Add('#');
			return new StandardValuesCollection(list);
		}

		/// <summary>
		/// If returns true it's not possible to type any value to the edit field. The attached property can only be changed via drop-down list.
		/// </summary>
		/// <remarks>This function is called by the framework on Main Thread</remarks>
		public override bool GetStandardValuesExclusive(SCM.ITypeDescriptorContext context)
		{
			return true;
		}

		/// <summary>
		/// A drop-down list is created for the attached property, if this returns true
		/// </summary>
		/// <remarks>This function is called by the framework on Main Thread</remarks>
		public override bool GetStandardValuesSupported(SCM.ITypeDescriptorContext context)
		{
			return true;
		}
	}

	internal class EmbedFigureOptionPageGrid : MVSS.DialogPage
	{
		/// <summary>
		/// The properties in <see cref="m_Options"/> will be listed in the options grid. 
		/// </summary>
		/// <remarks>
		/// <see cref="AutomationObject"/> returns with <see cref="m_Options"/>.
		/// </remarks>
		private EmbedFigureOptions m_Options = EmbedFigureOptions.Instance;

		/// <summary>
		/// Properties that are specified in <see cref="AutomationObject"/> are listed in the options grid.
		/// </summary>
		/// <remarks>If <see cref="AutomationObject"/> were not overridden, it would return with this.</remarks>
		public override object AutomationObject => m_Options;

		public override void LoadSettingsFromStorage()
		{
#if TRACE_FUNCTIONS
			Trace.EnterFunction();
#endif
			m_Options.LoadSettingsFromStorage();
#if TRACE_FUNCTIONS
			Trace.LeaveFunction();
#endif
		}

		public override void LoadSettingsFromXml(MVSSI.IVsSettingsReader reader)
		{
#if TRACE_FUNCTIONS
			Trace.EnterFunction();
#endif
			m_Options.LoadSettingsFromXml(reader);
#if TRACE_FUNCTIONS
			Trace.LeaveFunction();
#endif
		}


		public override void SaveSettingsToStorage()
		{
#if TRACE_FUNCTIONS
			Trace.EnterFunction();
#endif
			m_Options.SaveSettingsToStorage();
#if TRACE_FUNCTIONS
			Trace.LeaveFunction();
#endif
		}

		public override void SaveSettingsToXml(MVSSI.IVsSettingsWriter writer)
		{
#if TRACE_FUNCTIONS
			Trace.EnterFunction();
#endif
			m_Options.SaveSettingsToXml(writer);
#if TRACE_FUNCTIONS
			Trace.LeaveFunction();
#endif
		}

		public override void ResetSettings()
		{
#if TRACE_FUNCTIONS
			Trace.EnterFunction();
#endif
			m_Options.ResetSettings();
#if TRACE_FUNCTIONS
			Trace.LeaveFunction();
#endif
		}
	}

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
	// This specifies info that's displayed in About window
	[MVSS.InstalledProductRegistration("EmbedFigure", "Embed figures into your source code.", "0.1.0.0")]
	[MVSS.PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
	// This is for Tools -> Options
	[MVSS.ProvideOptionPage(typeof(EmbedFigureOptionPageGrid), "EmbedFigure_ProvideOptionPage_CaregoryName", "EmbedFigure_ProvideOptionPage_ObjectName", 1001, 1002, true)]
	// This is for Tools -> Import and Export Settings...
	// categoryName and objectName are save into the XML file concatenated with an "_" as name and RegisteredName attributes in Category tag
	// obectNameResourceID shows up in Import and Export Settings...
	[MVSS.ProvideProfile(typeof(EmbedFigureOptionPageGrid), "EmbedFigure", "General", 1003, 1004, false)]
	[SRIS.Guid(c_PackageGuidString)]
	internal sealed class EmbedFigurePackage : MVSS.AsyncPackage
	{
		/// <summary>
		/// EmbedFigurePackage GUID string.
		/// </summary>
		public const string c_PackageGuidString = "00F0521E-F06B-4187-A993-F3038B380531";

		static EmbedFigurePackage()
		{
			// Inside this method you can place any initialization code that does not require
			// any Visual Studio service because at this point the package object is created but
			// not sited yet inside Visual Studio environment. The place to do all the other
			// initialization is the Initialize method.
		}


		/// <summary>
		/// Initializes a new instance of the <see cref="EmbedFigurePackage"/> class.
		/// </summary>
		/// <remarks>
		/// This needs to be public, otherwise the options won't work, because Visual Studio throws a <see cref="S.MissingMethodException"/> and can't load the package.
		/// </remarks>
		public EmbedFigurePackage()
		{
			// Inside this method you can place any initialization code that does not require
			// any Visual Studio service because at this point the package object is created but
			// not sited yet inside Visual Studio environment. The place to do all the other
			// initialization is the Initialize method.
		}

		#region Package Members

		/// <summary>
		/// Initialization of the package; this method is called right after the package is sited, so this is the place
		/// where you can put all the initialization code that rely on services provided by VisualStudio.
		/// </summary>
		/// <param name="cancellationToken">A cancellation token to monitor for initialization cancellation, which can occur when VS is shutting down.</param>
		/// <param name="progress">A provider for progress updates.</param>
		/// <returns>A task representing the async work of package initialization, or an already completed task if there is none. Do not return null from this method.</returns>
		protected override async STT.Task InitializeAsync(ST.CancellationToken cancellationToken, S.IProgress<MVSS.ServiceProgressData> progress)
		{
			// When initialized asynchronously, the current thread may be a background thread at this point.
			// Do any initialization that requires the UI thread after switching to the UI thread.
			await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
		}

		#endregion
	}
}
