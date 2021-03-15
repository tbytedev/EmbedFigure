/*
 * EmbedFigure - Visual Studio extension for embedding math figures into source code
 * Copyright(C) 2021 Tamas Kezdi
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

//#undef DEBUG
#undef TRACE

#if TRACE
#define TRACE_LAYOUT_CHANGED
#define TRACE_ADORNMENT_ADD_LINE_NUMBER
#define TRACE_ADORNMENT_REMOVE_LINE_NUMBER
#define TRACE_LINE_TRANSFORM_LINE_NUMBERS
#endif

#define HANDLE_FOCUS

using MVS   = Microsoft.VisualStudio;
using MVSS  = Microsoft.VisualStudio.Shell;
using MVSSI = Microsoft.VisualStudio.Shell.Interop;
using MVST  = Microsoft.VisualStudio.Text;
using MVSTE = Microsoft.VisualStudio.Text.Editor;
using MVSTF = Microsoft.VisualStudio.Text.Formatting;
using S     = System;
using SCG   = System.Collections.Generic;
using SD    = System.Drawing;
using SDI   = System.Drawing.Imaging;
using SIO   = System.IO;
using ST    = System.Timers;
using STT   = System.Threading.Tasks;
using SW    = System.Windows;
using SWC   = System.Windows.Controls;
using SWM   = System.Windows.Media;
using SWMI  = System.Windows.Media.Imaging;

#if TRACE || DEBUG
using TRC_SD   = System.Diagnostics;
#endif

#if TRACE
using TRC_ST   = System.Threading;
using TRC_SRCS = System.Runtime.CompilerServices;
#endif

namespace EmbedFigure
{
	internal enum ColorTheme
	{
		Unspecified,
		Light,
		Dark
	}

	internal struct ChangeRanges
	{
		internal int m_OldFirstLineNumber;
		internal int m_OldLastLineNumber;
		internal int m_NewFirstLineNumber;
		internal int m_NewLastLineNumber;
		internal int m_LineCountDelta;
	}

	internal struct LineNumberShift
	{
		internal int m_OldLineNumber;
		internal int m_NewLineNumber;
		internal LineEntry m_LineEntry;
	}

	/// <summary>
	/// Contains the rendered figure, that is ready to add to the <see cref="MVSTE.IAdornmentLayer">IAdornmentLayer</see>
	/// </summary>
	internal class Figure : S.IDisposable
	{
		internal readonly SWMI.BitmapImage m_BitmapImage;
		internal readonly double m_Height;

		private readonly SIO.MemoryStream m_MemoryStream;

		internal Figure(SIO.MemoryStream memory_stream, double height)
		{
			// Load bitmap to System.Windows.Media.Imaging.BitmapImage from System.IO.MemoryStream
			m_BitmapImage = new SWMI.BitmapImage();
			m_BitmapImage.BeginInit();
			m_BitmapImage.StreamSource = memory_stream;
			m_BitmapImage.EndInit();

			m_MemoryStream = memory_stream;
			m_Height = height;
		}

		public void Dispose()
		{
			m_MemoryStream.Dispose();
		}
	}

	internal class FigureCacheData
	{
		internal readonly SCG.HashSet<LineID> m_LineIDs;

		internal Figure m_Figure;
		internal uint m_UpdateID = uint.MaxValue;
		internal S.DateTime m_LastWriteTimeUtc = S.DateTime.MinValue;

		internal FigureCacheData(LineID line_id)
		{
			m_LineIDs = new SCG.HashSet<LineID>() { line_id };
		}

		internal void AddLineID(LineID line_id)
		{
			m_LineIDs.Add(line_id);
		}

		internal bool RemoveLineID(LineID line_id)
		{
			m_LineIDs.Remove(line_id);
			return 0 == m_LineIDs.Count;
		}
	}

	internal class FigureCacheEntry
	{
		internal readonly string m_FigurePath;
		internal readonly double m_ZoomLevel;
		internal readonly bool m_Inverted;

		internal FigureCacheData m_FigureCacheData;

		internal FigureCacheEntry(string figure_path, double zoom_level, bool inverted)
		{
			m_FigurePath = figure_path;
			m_ZoomLevel = zoom_level;
			m_Inverted = inverted;
		}

		/// <summary>
		/// Determines whether the two <see cref="FigureCacheEntry"/> instances are identical for a <see cref="SCG.HashSet{T}">HashSet</see>
		/// </summary>
		/// <remarks>
		/// <para><see cref="FigureCacheEntry"/> is used in <see cref="SCG.HashSet{T}">HashSet</see>,
		/// which imitates the behavior of a <see cref="SCG.Dictionary{TKey, TValue}">Dictionaty</see>.
		/// <see cref="FigureCacheData"/> stores the value and the key consists of the remaining members.
		/// <see cref="SCG.Dictionary{TKey, TValue}.TryGetValue(TKey, out TValue)">Dictionary.TryGetValue</see>
		/// can't give back the stored key, but <see cref="SCG.HashSet{T}.TryGetValue(T, out T)">HashSet.TryGetValue</see>
		/// can give back the the entire stored value, from a partially set up value which only contains the key members.</para>
		/// <para>Comparing by reference is still possible via operator ==</para>
		/// </remarks>
		public override bool Equals(object obj)
		{
			return obj is FigureCacheEntry entry
				&& m_FigurePath == entry.m_FigurePath
				&& m_ZoomLevel  == entry.m_ZoomLevel
				&& m_Inverted   == entry.m_Inverted;
		}

		/// <summary>
		/// Hash for <see cref="FigureCacheEntry"/>
		/// </summary>
		/// <remarks>
		/// <para><see cref="FigureCacheEntry"/> is used in <see cref="SCG.HashSet{T}">HashSet</see>,
		/// which imitates the behavior of a <see cref="SCG.Dictionary{TKey, TValue}">Dictionaty</see>.
		/// <see cref="FigureCacheData"/> stores the value and the key consists of the remaining members.
		/// <see cref="SCG.Dictionary{TKey, TValue}.TryGetValue(TKey, out TValue)">Dictionary.TryGetValue</see>
		/// can't give back the stored key, but <see cref="SCG.HashSet{T}.TryGetValue(T, out T)">HashSet.TryGetValue</see>
		/// can give back the the entire stored value, from a partially set up value which only contains the key members.</para>
		/// <para>Generated by Visual Studio.</para>
		/// </remarks>
		public override int GetHashCode()
		{
			int hashCode = 2138992742;
			hashCode = hashCode * -1521134295 + SCG.EqualityComparer<string>.Default.GetHashCode(m_FigurePath);
			hashCode = hashCode * -1521134295 + m_ZoomLevel.GetHashCode();
			hashCode = hashCode * -1521134295 + m_Inverted.GetHashCode();
			return hashCode;
		}

	}

	internal readonly struct FigureLoadQueueEntry
	{
		internal readonly EmbedFigureManager m_Manager;
		internal readonly FigureCacheEntry m_FigureCacheEntry;
		internal readonly ColorTheme m_ColorTheme;

		internal FigureLoadQueueEntry(EmbedFigureManager manager, FigureCacheEntry figure_cache_entry, ColorTheme color_tone)
		{
			m_Manager = manager;
			m_FigureCacheEntry = figure_cache_entry;
			m_ColorTheme = color_tone;
		}
	}

	internal class LineEntry
	{
		internal FigureCacheEntry m_FigureCacheEntry;
		internal Figure m_Figure;
		internal ColorTheme m_ColorTheme;
		internal bool m_Added = false;

		internal LineEntry(ColorTheme color_theme)
		{
			m_ColorTheme = color_theme;
		}
	}

	internal readonly struct LineID
	{
		internal readonly EmbedFigureManager m_Manager;
		internal readonly int m_LineNumber;

		internal LineID(EmbedFigureManager manager, int line_number)
		{
			m_Manager = manager;
			m_LineNumber = line_number;
		}
	}

	internal enum ParameterType
	{
		UnknownParameter,
		ColorTheme,
		SVGFile,
		ParameterValue
	}

	internal readonly struct ParameterToken
	{
		internal readonly string m_RawText;
		internal readonly ParameterType m_Type;

		internal ParameterToken(string raw_text, ParameterType type)
		{
			m_RawText = raw_text;
			m_Type = type;
		}
	}

#if HANDLE_FOCUS
	internal class VSEvents : MVSSI.IVsBroadcastMessageEvents
	{
		const uint WM_ACTIVATEAPP = 0x001C;

		/// <summary>
		/// Fires when a message is broadcast to the environment window.
		/// </summary>
		/// <remarks>Invalidates figures when user switches to another application by increasing <see cref="EmbedFigureManager.s_UpdateID"/>.</remarks>
		/// <param name="msg">Specifies the notification message.</param>
		/// <param name="w_param">Word value parameter for the Windows message, as received by the environment.</param>
		/// <param name="l_param">Long integer parameter for the Windows message, as received by the environment.</param>
		/// <returns>If the method succeeds, it returns <see cref="MVS.VSConstants.S_OK"/>. If it fails, it returns an error code.</returns>
		public int OnBroadcastMessage(uint msg, S.IntPtr w_param, S.IntPtr l_param)
		{
			if (WM_ACTIVATEAPP == msg && S.IntPtr.Zero == w_param)
			{
				++EmbedFigureManager.s_UpdateID;
			}
			return MVS.VSConstants.S_OK;
		}
	}
#endif

	/// <summary>
	/// TextAdornment to place figures after #EmbedFigure instructions
	/// </summary>
	internal class EmbedFigureManager
	{
		// ----------------------------------------------------------------
		// Static members
		// ----------------------------------------------------------------

		/// <summary>
		/// Stores rendered figures for each path in <see cref="SWMI.BitmapImage">BitmapImages</see>
		/// It's accessed only from Main thread
		/// </summary>
		internal static readonly SCG.HashSet<FigureCacheEntry> s_FigureCache = new SCG.HashSet<FigureCacheEntry>();

		internal static readonly SCG.List<EmbedFigureManager> s_Managers = new SCG.List<EmbedFigureManager>();

		/// <summary>
		/// Instruction to embed a figure to the source. This is prefixed by a #
		/// </summary>
		private static readonly char[] s_InstructionCharArray = { 'E', 'm', 'b', 'e', 'd', 'F', 'i', 'g', 'u', 'r', 'e' };

		private static readonly ParameterToken[] s_ParameterDefinitions =
		{
			new ParameterToken("ColorTheme", ParameterType.ColorTheme),
			new ParameterToken("SVGFile", ParameterType.SVGFile)
		};

		/// <summary>
		/// List of characters, that are not allowed in file names
		/// </summary>
		private static readonly char[] s_InvalidChars;

		/// <summary>
		/// This timer is fired after the user hasn't changed the text for 1500 ms and there are files to load, or the cache can be cleaned up.
		/// </summary>
		/// <remarks>
		/// Do not load and render figures at once while user is still typing, rather wait some time to let things settle down a bit.
		/// Load is commenced when this timer is elapsed.
		/// </remarks>
		private static readonly ST.Timer s_Timer = new ST.Timer(1500);

		/// <summary>
		/// Stores the figures to load and the lines to refresh
		/// It's accessed only from Main thread
		/// </summary>
		private static readonly SCG.HashSet<FigureLoadQueueEntry> s_FigureLoadQueue = new SCG.HashSet<FigureLoadQueueEntry>();

#if HANDLE_FOCUS
		private static readonly MVSSI.IVsShell s_VSShell;
		private static readonly VSEvents s_VSEvents;
#endif

#if TRACE
		private static readonly TRC_ST.ThreadLocal<string> s_ThreadName = new TRC_ST.ThreadLocal<string>(() => { return "Thread " + (10 > TRC_ST.Thread.CurrentThread.ManagedThreadId ? " " + TRC_ST.Thread.CurrentThread.ManagedThreadId : TRC_ST.Thread.CurrentThread.ManagedThreadId.ToString()); });
		private static readonly TRC_ST.ThreadLocal<int>    s_Indent     = new TRC_ST.ThreadLocal<int>   (() => { return 0; });
#endif

		internal static uint s_UpdateID = 0;

		private static bool s_CacheCanHaveUnreferencedEntries = false;

		private static ST.ElapsedEventHandler s_TimerEventHandler;

		/// <summary>
		/// Indicates if the timer is active.
		/// </summary>
		/// <remarks>
		/// <see cref="TimerElapsed"/> may be raised even after <see cref="s_Timer"/> has been stopped.
		/// The actual value of this ID is copied into the timer event handler when <see cref="StartTimer"/> is called, and this ID is incremented when <see cref="StopTimer"/> is called.
		/// So when the event is raised the event handler can compare its copied ID and the actual ID.
		/// If the two values are the same no Stop() method has been called between the timer start and timer raise.
		/// </remarks>
		private static int s_TimerStartID = 0;

		/// <summary>
		/// Set up a sorted list of invalid characters, because it's faster to search in sorted arrays.
		/// </summary>
		static EmbedFigureManager()
		{
			var invalid_chars = new SCG.List<char>(SIO.Path.GetInvalidFileNameChars());

			// Remove those characters from invalid list which are valid in paths
			invalid_chars.Remove(SIO.Path.DirectorySeparatorChar);
			invalid_chars.Remove(SIO.Path.AltDirectorySeparatorChar);
			invalid_chars.Remove(SIO.Path.VolumeSeparatorChar);

			s_InvalidChars = invalid_chars.ToArray();
			S.Array.Sort(s_InvalidChars);

			s_Timer.AutoReset = false;

#if HANDLE_FOCUS
			// MVSSI.IVsShell can only be used on main thread. Compiler generates warning if ThrowIfNotOnUIThread is not called before using MVSSI.IVsShell
			MVSS.ThreadHelper.ThrowIfNotOnUIThread();
			s_VSShell = MVSS.Package.GetGlobalService(typeof(MVSSI.SVsShell)) as MVSSI.IVsShell;
			if (null != s_VSShell)
			{
				s_VSEvents = new VSEvents();
				s_VSShell.AdviseBroadcastMessages(s_VSEvents, out uint cookie);
			}
#endif
		}

		private static void CacheCleanup()
		{
			var cache_entries_to_remove = new SCG.List<FigureCacheEntry>();

			foreach (FigureCacheEntry figure_cache_entry in s_FigureCache)
			{
				FigureCacheData figure_cache_data = figure_cache_entry.m_FigureCacheData;
				if (0 == figure_cache_data.m_LineIDs.Count)
				{
					cache_entries_to_remove.Add(figure_cache_entry);
					UnloadFigure(figure_cache_data);
				}
			}

			foreach (FigureCacheEntry figure_cache_entry in cache_entries_to_remove)
			{
				s_FigureCache.Remove(figure_cache_entry);
			}

			s_CacheCanHaveUnreferencedEntries = false;
		}

		private static ColorTheme GetColorThemeFromBrush(SWM.Brush brush)
		{
			if (brush is SWM.SolidColorBrush solid_color_brush)
			{
				float luminosity = 0.2126f * solid_color_brush.Color.ScR + 0.7152f * solid_color_brush.Color.ScG + 0.0722f * solid_color_brush.Color.ScB;
				if (0.5 < luminosity)
				{
					return ColorTheme.Light;
				}
				else
				{
					return ColorTheme.Dark;
				}
			}
			return ColorTheme.Unspecified;
		}

		private static void StartTimer()
		{
			// Capture current value of s_TimerStartID to be used in event handler lambda expression.
			int timer_start_id = s_TimerStartID;

			// Create new delegate from lambda with captured timer start ID, and store it to s_TimerEventHandler to be able to unsubscribe later.
			s_TimerEventHandler = (sender, e) => { TimerElapsed(timer_start_id); };

			// It's necessary to subscribe a new delegate each time the timer has been started, because every delegate contains a different captured timer start id.
			s_Timer.Elapsed += s_TimerEventHandler;
			s_Timer.Start();
		}

		private static void StopTimer()
		{
			// It's possible that Elapsed event is raised after the Stop method is called. Increasing timer start ID invalidates upcoming elapsed event.
			++s_TimerStartID;
			s_Timer.Stop();

			// Unsubscribe event handler
			if (null != s_TimerEventHandler)
			{
				s_Timer.Elapsed -= s_TimerEventHandler;
				s_TimerEventHandler = null;
			}
		}

		public static void LoadFigureTask(object context)
		{
#if TRACE
			EnterFunction();
#endif
			var figure_cache_entry = (FigureCacheEntry)context;
			SD.Bitmap bitmap = null;
			SIO.MemoryStream memory_stream = null;
			double height = 0.0;
			try
			{
				Svg.SvgDocument svg_doc = Svg.SvgDocument.Open(figure_cache_entry.m_FigurePath);
				bitmap = svg_doc.Draw();
				height = bitmap.Height;
				if (100.0 != figure_cache_entry.m_ZoomLevel)
				{
					double zoom = figure_cache_entry.m_ZoomLevel / 100.0;
					bitmap = svg_doc.Draw(S.Convert.ToInt32(bitmap.Width * zoom), S.Convert.ToInt32(bitmap.Height * zoom));
				}

				// Invert figure if needed
				if (figure_cache_entry.m_Inverted)
				{
					for (int x = 0; x < bitmap.Width; ++x)
					{
						for (int y = 0; y < bitmap.Height; ++y)
						{
							SD.Color original_color = bitmap.GetPixel(x, y);
							SD.Color new_color = SD.Color.FromArgb(original_color.A, 255 - original_color.R, 255 - original_color.G, 255 - original_color.B);
							bitmap.SetPixel(x, y, new_color);
						}
					}
				}

				// Convert System.Drawing.Bitmap to System.Windows.Controls.Image, that can be added to m_AdornmentLayer
				// Save System.Drawing.Bitmap to a System.IO.MemoryStream
				memory_stream = new SIO.MemoryStream();
				bitmap.Save(memory_stream, SDI.ImageFormat.Tiff);
				memory_stream.Seek(0, SIO.SeekOrigin.Begin);

#if TRACE
				TraceMsg("Switch to Main LoadFigureTask");
#endif
				// UI related objects (System.Windows.Media.Imaging.BitmapImage) can be created and used only on Main thread
				MVSS.ThreadHelper.JoinableTaskFactory.Run(async delegate
				{
					await MVSS.ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

#if TRACE
					TraceMsg("Switched to Main LoadFigureTask");
#endif
					FigureCacheData figure_cache_data = figure_cache_entry.m_FigureCacheData;
					if (0 == figure_cache_data.m_LineIDs.Count)
					{
						// This figure has been removed after the load process was started.
						memory_stream.Dispose();
						return;
					}

					figure_cache_data.m_Figure = new Figure(memory_stream, height);

					foreach (LineID line_id in figure_cache_data.m_LineIDs)
					{
						EmbedFigureManager manager = line_id.m_Manager;
						int line_number = line_id.m_LineNumber;
						LineEntry line_entry = manager.m_LineEntries[line_number];

						manager.AddFigure(line_number, line_entry, figure_cache_data);
					}
#if TRACE
					TraceMsg("Switch from Main LoadFigureTask");
#endif
				});
			}
			catch
			{ }
			finally
			{
#if TRACE
				TraceMsg("Switched from Main LoadFigureTask");
#endif
				bitmap?.Dispose();
			}

#if TRACE
			LeaveFunction();
#endif
			return;
		}

		private static void ProcessLineLoadQueue()
		{
			// Create a list of lines for each figure. It's possible that more than one lines are waiting for the same figure to be loaded.
			var figure_load_queue = new SCG.Dictionary<FigureCacheEntry, SCG.HashSet<LineID>>();
			foreach (FigureLoadQueueEntry figure_load_queue_entry in s_FigureLoadQueue)
			{
				FigureCacheEntry figure_cache_entry = figure_load_queue_entry.m_FigureCacheEntry;
				FigureCacheData figure_cache_data = figure_cache_entry.m_FigureCacheData;

				// Check if there are still lines referring to this cache entry
				if (0 == figure_cache_data.m_LineIDs.Count)
				{
					continue;
				}

				// Check if update is necessary
				if (figure_cache_data.m_UpdateID == s_UpdateID)
				{
					continue;
				}

				figure_cache_data.m_UpdateID = s_UpdateID;

				var file_info = new SIO.FileInfo(figure_cache_entry.m_FigurePath);

				if (!file_info.Exists)
				{
					// This figure has been deleted.
					UnloadFigure(figure_cache_data);
					continue;
				}
				if (file_info.LastWriteTimeUtc == figure_cache_data.m_LastWriteTimeUtc)
				{
					continue;
				}

				UnloadFigure(figure_cache_data);
				figure_cache_data.m_LastWriteTimeUtc = file_info.LastWriteTimeUtc;

				var task = new STT.Task(LoadFigureTask, figure_load_queue_entry.m_FigureCacheEntry);
				task.Start();
				// Although System.Threading.Tasks.Task is IDisposable, it's not necessary to call its Dispose() function.
				// https://devblogs.microsoft.com/pfxteam/do-i-need-to-dispose-of-tasks/
			}

			s_FigureLoadQueue.Clear();

		}

		private static void UnloadFigure(FigureCacheData figure_cache_data)
		{
			if (null != figure_cache_data.m_Figure)
			{
				foreach (LineID line_id in figure_cache_data.m_LineIDs)
				{
					EmbedFigureManager manager = line_id.m_Manager;
					int line_number = line_id.m_LineNumber;
					LineEntry line_entry = manager.m_LineEntries[line_number];
					if (line_entry.m_Added)
					{
						manager.m_AdornmentLayer.RemoveAdornmentsByTag(line_number);
					}
					line_entry.m_Figure = null;
				}
				figure_cache_data.m_Figure.Dispose();
				figure_cache_data.m_Figure = null;
				figure_cache_data.m_LastWriteTimeUtc = S.DateTime.MinValue;
			}
		}

		/// <summary>
		/// This timer is fired after the user hasn't changed the text for 1500 ms and there are files to load.
		/// </summary>
		/// <remarks>This function is called by the framework on a Worker Thread</remarks>
		private static void TimerElapsed(int timer_start_id)
		{
#if TRACE
			EnterFunction();
			TraceMsg("Switch to Main TimerElapsed");
#endif
			MVSS.ThreadHelper.JoinableTaskFactory.Run(async delegate
			{
				await MVSS.ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
				if (s_TimerStartID != timer_start_id)
				{
					return;
				}

#if TRACE
				TraceMsg("Switched to Main TimerElapsed");
#endif
				ProcessLineLoadQueue();
				CacheCleanup();
#if TRACE
				TraceMsg("Switch from Main TimerElapsed");
#endif
			});
#if TRACE
			TraceMsg("Switched from Main TimerElapsed");
#endif
#if TRACE
			LeaveFunction();
#endif
		}

		// ----------------------------------------------------------------
		// Non static members
		// ----------------------------------------------------------------

		/// <summary>
		/// The layer of the adornment.
		/// </summary>
		private readonly MVSTE.IAdornmentLayer m_AdornmentLayer;

		/// <summary>
		/// Text document. Needed for retrieving file path.
		/// </summary>
		private readonly MVST.ITextDocument m_TextDocument;

		/// <summary>
		/// Text view where the adornment is created.
		/// </summary>
		internal readonly MVSTE.IWpfTextView m_TextView;

		/// <summary>
		/// Stores figure info for each line in this manager
		/// </summary>
		internal SCG.SortedDictionary<int, LineEntry> m_LineEntries = new SCG.SortedDictionary<int, LineEntry>();

		/// <summary>
		/// Current color theme
		/// </summary>
		private ColorTheme m_ColorTheme = ColorTheme.Unspecified;

		/// <summary>
		/// Initializes a new instance of the <see cref="EmbedFigureManager"/> class.
		/// </summary>
		/// <param name="text_view">Text view to create the adornment for</param>
		internal EmbedFigureManager(MVSTE.IWpfTextView text_view, MVST.ITextDocumentFactoryService text_document_factory_service)
		{
			if (text_view == null)
			{
				throw new S.ArgumentNullException("TextView");
			}

			if (!text_document_factory_service.TryGetTextDocument(text_view.TextBuffer, out m_TextDocument))
			{
				// Document doesn't exist for textView.TextBuffer
				return;
			}

			s_Managers.Add(this);

			m_AdornmentLayer = text_view.GetAdornmentLayer("EmbedFigureAdornmentLayer");

			m_TextView = text_view;

			// Detect background color brightness
			m_ColorTheme = GetColorThemeFromBrush(m_TextView.Background);

			m_TextView.BackgroundBrushChanged += OnBackgroundChanged;
			m_TextView.Closed                 += OnClosed;
#if HANDLE_FOCUS
			m_TextView.GotAggregateFocus      += OnGotFocus;
#endif
			m_TextView.LayoutChanged          += OnLayoutChanged;
			m_TextView.ZoomLevelChanged       += OnZoomLevelChanged;
		}

		private void AddAdornment(MVST.SnapshotSpan snapshot_span, int line_number, LineEntry line_entry)
		{
#if TRACE
			EnterFunction();
#endif
			SWM.Geometry geometry = m_TextView.TextViewLines.GetMarkerGeometry(snapshot_span);
			if (null != geometry)
			{
				var image = new SWC.Image
				{
					Source = line_entry.m_Figure.m_BitmapImage,
					Height = line_entry.m_Figure.m_Height
				};

				SWC.Canvas.SetLeft(image, geometry.Bounds.Left);
				SWC.Canvas.SetTop(image, geometry.Bounds.Bottom);
				m_AdornmentLayer.AddAdornment(MVSTE.AdornmentPositioningBehavior.TextRelative, snapshot_span, line_number, image, OnAdornmentRemoved);
				line_entry.m_Added = true;
#if TRACE_ADORNMENT_ADD_LINE_NUMBER
				TraceMsg("Adornment add line number: " + line_number);
#endif
			}
#if TRACE
			LeaveFunction();
#endif
		}

		private void AddVisibleAdornments(int first_line_number_to_add_adornment, int last_line_number_to_add_adornment)
		{
			int first_visible_line_number = m_TextView.TextSnapshot.GetLineNumberFromPosition(m_TextView.TextViewLines.FormattedSpan.Start);
			int last_visible_line_number = m_TextView.TextSnapshot.GetLineNumberFromPosition(m_TextView.TextViewLines.FormattedSpan.End - 1);

			foreach (SCG.KeyValuePair<int, LineEntry> pair in m_LineEntries)
			{
				int line_number = pair.Key;
				if (line_number < first_visible_line_number)
				{
					continue;
				}
				if (line_number < first_line_number_to_add_adornment)
				{
					continue;
				}
				if (last_visible_line_number < line_number)
				{
					break;
				}
				if (last_line_number_to_add_adornment < line_number)
				{
					break;
				}

				LineEntry line_entry = pair.Value;
				if (null == line_entry.m_Figure || line_entry.m_Added)
				{
					continue;
				}
				MVST.ITextSnapshotLine text_snapshot_line = m_TextView.TextSnapshot.GetLineFromLineNumber(line_number);
				AddAdornment(text_snapshot_line.Extent, line_number, line_entry);
			}
		}

		private void AddFigure(int line_number, LineEntry line_entry, FigureCacheData figure_cache_data)
		{
			line_entry.m_Figure = figure_cache_data.m_Figure;
			if (null == line_entry.m_Figure)
			{
				return;
			}

			// Get the first line from m_TextView.TextViewLines, this is not the same as m_TextView.TextViewLines.FirstVisibleLine.
			// It's possible that m_TextView.TextViewLines[0] is hidden and it's before m_TextView.TextViewLines.FirstVisibleLine
			MVSTF.IWpfTextViewLine first_view_line = m_TextView.TextViewLines[0];
			int first_view_line_number = m_TextView.TextSnapshot.GetLineNumberFromPosition(first_view_line.Start);
			int view_line_number = line_number - first_view_line_number;

			// Skip lines that has no corresponding line in m_TextView.TextViewLines
			if (0 > view_line_number)
			{
				return;
			}
			if (m_TextView.TextViewLines.Count <= view_line_number)
			{
				return;
			}

			// Refresh line
			MVSTF.IWpfTextViewLine curr_view_line = m_TextView.TextViewLines[view_line_number];
			AddAdornment(curr_view_line.Extent, line_number, line_entry);
			// Force calling of GetLineTransform for this line
			m_TextView.DisplayTextLineContainingBufferPosition(curr_view_line.Start, curr_view_line.Top - m_TextView.ViewportTop, MVSTE.ViewRelativePosition.Top);
		}

		private SCG.List<ParameterToken> TokenizeParameters(string parameters_string)
		{
			var tokens = new SCG.List<ParameterToken>();
			int token_start_index = 0;
			int parameters_string_lenght = parameters_string.Length;

			int current_index = 0;
			for (;;)
			{
				if ('"' == parameters_string[current_index])
				{
					// Parameter is surrounded by "s
					int text_start_index = token_start_index + 1;
					++current_index;
					if (current_index == parameters_string_lenght)
					{
						return tokens;
					}

					bool quote_escaped = false;
					string unescaped_string = null;
					for (;;)
					{
						// Look for closing "
						while ('"' != parameters_string[current_index])
						{
							++current_index;
							if (current_index == parameters_string_lenght)
							{
								if (quote_escaped)
								{
									tokens.Add(new ParameterToken(unescaped_string + parameters_string.Substring(text_start_index, current_index - text_start_index), ParameterType.ParameterValue));
								}
								else
								{
									tokens.Add(new ParameterToken(parameters_string.Substring(text_start_index, current_index - text_start_index), ParameterType.ParameterValue));
								}
								return tokens;
							}
						}

						if ('\\' != parameters_string[current_index - 1])
						{
							// This isn't a closing ", because it's escaped by the preceding \
							if (quote_escaped)
							{
								tokens.Add(new ParameterToken(unescaped_string + parameters_string.Substring(text_start_index, current_index - text_start_index), ParameterType.ParameterValue));
							}
							else
							{
								tokens.Add(new ParameterToken(parameters_string.Substring(text_start_index, current_index - text_start_index), ParameterType.ParameterValue));
							}
							break;
						}

						// We need a new string in which the escaping \s are excluded. Substring of the parameter_string can't be used.
						if (!quote_escaped)
						{
							quote_escaped = true;
							unescaped_string = parameters_string.Substring(text_start_index, current_index - text_start_index - 1);
						}
						else
						{
							unescaped_string += parameters_string.Substring(text_start_index, current_index - text_start_index - 1);
						}
						text_start_index = current_index;

						// Advance to the next character after escaped "
						++current_index;
						if (current_index == parameters_string_lenght)
						{
							if (quote_escaped)
							{
								tokens.Add(new ParameterToken(unescaped_string + parameters_string.Substring(text_start_index, current_index - text_start_index), ParameterType.ParameterValue));
							}
							else
							{
								tokens.Add(new ParameterToken(parameters_string.Substring(text_start_index, current_index - text_start_index), ParameterType.ParameterValue));
							}
							return tokens;
						}
					}

					// Advance to the next character after closing "
					++current_index;
					if (current_index == parameters_string_lenght)
					{
						return tokens;
					}

					// " must be followed by a space
					if (!char.IsWhiteSpace(parameters_string[current_index]))
					{
						return tokens;
					}

					// Skip spaces after "
					while (char.IsWhiteSpace(parameters_string[current_index]))
					{
						++current_index;
						if (current_index == parameters_string_lenght)
						{
							return tokens;
						}
					}
				}
				else
				{
					// Look for closing space or :
					while (!char.IsWhiteSpace(parameters_string[current_index]) && ':' != parameters_string[current_index])
					{
						++current_index;
						if (current_index == parameters_string_lenght)
						{
							tokens.Add(new ParameterToken(parameters_string.Substring(token_start_index, current_index - token_start_index), ParameterType.ParameterValue));
							return tokens;
						}
					}

					string parameter_name = parameters_string.Substring(token_start_index, current_index - token_start_index);

					// Skip spaces after parameter
					while (char.IsWhiteSpace(parameters_string[current_index]))
					{
						++current_index;
						if (current_index == parameters_string_lenght)
						{
							tokens.Add(new ParameterToken(parameter_name, ParameterType.ParameterValue));
							return tokens;
						}
					}

					if (':' == parameters_string[current_index])
					{
						// Parameter is followed by a :. It's a parameter name.
						// Look for supported parameters
						foreach (ParameterToken definition in s_ParameterDefinitions)
						{
							if (definition.m_RawText == parameter_name)
							{
								tokens.Add(definition);
								goto _parameter_found;
							}
						}
						tokens.Add(new ParameterToken(parameter_name, ParameterType.UnknownParameter));

					_parameter_found:

						// Advance to the next character after :
						++current_index;
						if (current_index == parameters_string_lenght)
						{
							return tokens;
						}

						// Skip spaces after :
						while (char.IsWhiteSpace(parameters_string[current_index]))
						{
							++current_index;
							if (current_index == parameters_string_lenght)
							{
								return tokens;
							}
						}
					}
					else
					{
						// Parameter isn't followed by a :. It's a parameter value
						tokens.Add(new ParameterToken(parameter_name, ParameterType.ParameterValue));
					}
				}

				token_start_index = current_index;
			}
		}

		/// <summary>
		/// Parse line. Search for #EmbedFigure instruction, and register figure changes.
		/// </summary>
		/// <param name="line">Line to add the adornments</param>
		private void ParseLine(string line_string, out string out_figure_path, out ColorTheme out_color_tone)
		{
			out_figure_path = null;
			out_color_tone = ColorTheme.Unspecified;
			int line_length = line_string.Length;

			int index_in_line;
			for (index_in_line = 0; index_in_line < line_length; ++index_in_line)
			{
			_again:
				if ('#' != line_string[index_in_line])
				{
					continue;
				}

				// '#' counts only if the previous character was not a letter or digit
				if (0 != index_in_line && char.IsLetterOrDigit(line_string[index_in_line - 1]))
				{
					continue;
				}

				++index_in_line;
				// Compare instruction
				for (int index_in_instruction = 0; index_in_instruction < s_InstructionCharArray.Length; ++index_in_instruction, ++index_in_line)
				{
					if (index_in_line == line_length)
					{
						return;
					}
					if (s_InstructionCharArray[index_in_instruction] != line_string[index_in_line])
					{
						goto _again;
					}
				}

				if (index_in_line == line_length)
				{
					return;
				}

				// Instruction must be followed by a whitespace
				if (!char.IsWhiteSpace(line_string[index_in_line]))
				{
					goto _again;
				}

				// Skip spaces between instruction and first parameter
				do
				{
					++index_in_line;
					if (index_in_line == line_length)
					{
						return;
					}
				}
				while (char.IsWhiteSpace(line_string[index_in_line]));

				SCG.List<ParameterToken> tokens = TokenizeParameters(line_string.Substring(index_in_line));

				ParameterType parameter_name = ParameterType.UnknownParameter;

				string local_path = null;

				foreach (ParameterToken token in tokens)
				{
					if (token.m_Type == ParameterType.ParameterValue)
					{
						switch (parameter_name)
						{
							case ParameterType.SVGFile:
								if (null == local_path)
								{
									if (0 == token.m_RawText.Length)
									{
										break;
									}

									// Local path mustn't contain invalid characters
									foreach (char c in token.m_RawText)
									{
										if (0 <= S.Array.BinarySearch(s_InvalidChars, c))
										{
											goto _invalid_local_path;
										}
									}

									char last_char = token.m_RawText[token.m_RawText.Length - 1];
									if (SIO.Path.DirectorySeparatorChar == last_char || SIO.Path.AltDirectorySeparatorChar == last_char)
									{
										// Last character is a directory separator. It can't be a filename, it's a directory.
										break;
									}

									local_path = token.m_RawText;
								}
							_invalid_local_path:
								break;
							case ParameterType.ColorTheme:
								if (ColorTheme.Unspecified == out_color_tone)
								{
									if ("dark" == token.m_RawText)
									{
										out_color_tone = ColorTheme.Dark;
									}
									else if ("light" == token.m_RawText)
									{
										out_color_tone = ColorTheme.Light;
									}
								}
								break;
						}
						parameter_name = ParameterType.UnknownParameter;
					}
					else
					{
						parameter_name = token.m_Type;
					}
				}

				if (null == local_path)
				{
					return;
				}
				// TODO: Check if file system is case sensitive.
				// For now we're assuming that it's a windows file system, so it's case insensitive.
				// Convert filename to lowercase to ignore character case.
				local_path = local_path.ToLowerInvariant();
				local_path = local_path.Replace(SIO.Path.AltDirectorySeparatorChar, SIO.Path.DirectorySeparatorChar);

				// TODO: Support for full path and other rooted path types
				out_figure_path = SIO.Path.GetDirectoryName(m_TextDocument.FilePath) + SIO.Path.DirectorySeparatorChar + local_path;
				return;
			}
		}

		/// <summary>
		/// Parse line. Search for #EmbedFigure instruction, and register figure changes.
		/// </summary>
		/// <param name="line">Line to add the adornments</param>
		private void ProcessLine(MVST.SnapshotSpan snapshot_span, string line_text, int line_number)
		{
			ParseLine(line_text, out string figure_path, out ColorTheme color_theme);

			var line_id = new LineID(this, line_number);

			if (null == figure_path)
			{
				// Currently there's no figure specified in this line
				if (m_LineEntries.TryGetValue(line_number, out LineEntry old_line_entry))
				{
					// But there was a figure in this line previously. Remove previous figure.
					RemoveFigure(line_number, old_line_entry);
					m_LineEntries.Remove(line_number);
				}
			}
			else
			{
				// There's a figure in this line
				double zoom_level = m_TextView.ZoomLevel;
				bool inverted = ColorTheme.Unspecified != color_theme && color_theme != m_ColorTheme;

				if (m_LineEntries.TryGetValue(line_number, out LineEntry line_entry))
				{
					// There's a figure in this line now but there was already a figure in this line
					line_entry.m_ColorTheme = color_theme;

					if (line_entry.m_FigureCacheEntry.m_FigurePath != figure_path ||
						line_entry.m_FigureCacheEntry.m_Inverted   != inverted ||
						line_entry.m_FigureCacheEntry.m_ZoomLevel  != zoom_level)
					{
						// The current and the previous figures are different
						RemoveFigure(line_number, line_entry);

						FigureCacheEntry figure_cache_entry = new FigureCacheEntry(figure_path, zoom_level, inverted);

						if (s_FigureCache.TryGetValue(figure_cache_entry, out FigureCacheEntry stored_figure_cache_entry))
						{
							// This figure is already loaded, so just use it.
							line_entry.m_FigureCacheEntry = stored_figure_cache_entry;
							FigureCacheData figure_cache_data = stored_figure_cache_entry.m_FigureCacheData;
							figure_cache_data.AddLineID(line_id);
							line_entry.m_Figure = figure_cache_data.m_Figure;
						}
						else
						{
							line_entry.m_FigureCacheEntry = figure_cache_entry;
							figure_cache_entry.m_FigureCacheData = new FigureCacheData(line_id);
							s_FigureCache.Add(figure_cache_entry);
							s_FigureLoadQueue.Add(new FigureLoadQueueEntry(this, figure_cache_entry, color_theme));
						}
					}
				}
				else
				{
					// There's a figure in this line and there was no figure in this line previously
					FigureCacheEntry figure_cache_entry = new FigureCacheEntry(figure_path, zoom_level, inverted);
					line_entry = new LineEntry(color_theme);
					if (s_FigureCache.TryGetValue(figure_cache_entry, out FigureCacheEntry stored_figure_cache_entry))
					{
						line_entry.m_FigureCacheEntry = stored_figure_cache_entry;
						FigureCacheData figure_cache_data = stored_figure_cache_entry.m_FigureCacheData;
						figure_cache_data.AddLineID(line_id);
						line_entry.m_Figure = figure_cache_data.m_Figure;
					}
					else
					{
						line_entry.m_FigureCacheEntry = figure_cache_entry;
						figure_cache_entry.m_FigureCacheData = new FigureCacheData(line_id);
						s_FigureCache.Add(figure_cache_entry);
						s_FigureLoadQueue.Add(new FigureLoadQueueEntry(this, figure_cache_entry, color_theme));
					}
					m_LineEntries.Add(line_number, line_entry);
				}
			}
		}

		private void RemoveAdornment(int line_number, LineEntry line_entry)
		{
			if (line_entry.m_Added)
			{
				m_AdornmentLayer.RemoveAdornmentsByTag(line_number);
			}
		}

		private void RemoveFigure(int line_number, LineEntry line_entry)
		{
			RemoveAdornment(line_number, line_entry);
			FigureCacheData figure_cache_data = line_entry.m_FigureCacheEntry.m_FigureCacheData;
			if (figure_cache_data.RemoveLineID(new LineID(this, line_number)))
			{
				s_CacheCanHaveUnreferencedEntries = true;
			}
			line_entry.m_Figure = null;
		}

		private void ShiftFigureLineNumber(LineNumberShift line_number_shift)
		{
			FigureCacheData figure_cache_data = line_number_shift.m_LineEntry.m_FigureCacheEntry.m_FigureCacheData;
			figure_cache_data.RemoveLineID(new LineID(this, line_number_shift.m_OldLineNumber));
			figure_cache_data.AddLineID(new LineID(this, line_number_shift.m_NewLineNumber));
		}

		private void OnAdornmentRemoved(object tag, SW.UIElement element)
		{
#if TRACE
			EnterFunction();
#endif
			int line_number = (int)tag;
			m_LineEntries[line_number].m_Added = false;
#if TRACE_ADORNMENT_REMOVE_LINE_NUMBER
			TraceMsg("Adornment remove line number: " + line_number);
			LeaveFunction();
#endif
		}

		private void OnBackgroundChanged(object sender, MVSTE.BackgroundBrushChangedEventArgs e)
		{
			ColorTheme color_theme = GetColorThemeFromBrush(e.NewBackgroundBrush);
			if (color_theme == m_ColorTheme)
			{
				return;
			}
#if TRACE
			EnterFunction();
#endif
			m_ColorTheme = color_theme;

			StopTimer();

			// Remove those lines from load queue, which are about to load with a different invert value
			var figure_load_entries_to_remove = new SCG.List<FigureLoadQueueEntry>();
			foreach (FigureLoadQueueEntry figure_load_queue_entry in s_FigureLoadQueue)
			{
				FigureCacheEntry figure_cache_entry = figure_load_queue_entry.m_FigureCacheEntry;
				ColorTheme load_entry_color_theme = figure_load_queue_entry.m_ColorTheme;
				bool inverted = ColorTheme.Unspecified != load_entry_color_theme && load_entry_color_theme != m_ColorTheme;
				if (this == figure_load_queue_entry.m_Manager && figure_cache_entry.m_Inverted != inverted)
				{
					figure_load_entries_to_remove.Add(figure_load_queue_entry);
				}
			}
			foreach (FigureLoadQueueEntry figure_load_queue_entry in figure_load_entries_to_remove)
			{
				s_FigureLoadQueue.Remove(figure_load_queue_entry);
			}

			foreach (SCG.KeyValuePair<int, LineEntry> pair in m_LineEntries)
			{
				LineEntry line_entry = pair.Value;
				int line_number = pair.Key;
				// If color theme is unspecified, leave the figure as it is
				if (ColorTheme.Unspecified == line_entry.m_ColorTheme)
				{
					continue;
				}

				RemoveFigure(line_number, line_entry);
				bool inverted = line_entry.m_ColorTheme != m_ColorTheme;
				var figure_cache_entry = new FigureCacheEntry(line_entry.m_FigureCacheEntry.m_FigurePath, line_entry.m_FigureCacheEntry.m_ZoomLevel, inverted);
				if (s_FigureCache.TryGetValue(figure_cache_entry, out FigureCacheEntry stored_figure_cache_entry))
				{
					line_entry.m_FigureCacheEntry = stored_figure_cache_entry;
					FigureCacheData figure_cache_data = stored_figure_cache_entry.m_FigureCacheData;
					figure_cache_data.AddLineID(new LineID(this, line_number));
					AddFigure(line_number, line_entry, figure_cache_data);
				}
				else
				{
					line_entry.m_FigureCacheEntry = figure_cache_entry;
					figure_cache_entry.m_FigureCacheData = new FigureCacheData(new LineID(this, line_number));
					s_FigureCache.Add(figure_cache_entry);
					s_FigureLoadQueue.Add(new FigureLoadQueueEntry(this, figure_cache_entry, line_entry.m_ColorTheme));
				}
			}

			ProcessLineLoadQueue();
			if (s_CacheCanHaveUnreferencedEntries)
			{
				StartTimer();
			}
#if TRACE
			LeaveFunction();
#endif
		}

		private void OnClosed(object sender, S.EventArgs e)
		{
			m_TextView.Properties.RemoveProperty(typeof(EmbedFigureManager));
			s_Managers.Remove(this);
		}

#if HANDLE_FOCUS
		private void OnGotFocus(object sender, S.EventArgs e)
		{
#if TRACE
			EnterFunction();
#endif
			StopTimer();

			s_FigureLoadQueue.Clear();

			foreach (SCG.KeyValuePair<int, LineEntry> pair in m_LineEntries)
			{
				LineEntry line_entry = pair.Value;
				int line_number = pair.Key;
				s_FigureLoadQueue.Add(new FigureLoadQueueEntry(this, line_entry.m_FigureCacheEntry, line_entry.m_ColorTheme));
			}

			ProcessLineLoadQueue();
#if TRACE
			LeaveFunction();
#endif
		}
#endif

		/// <summary>
		/// Handles whenever the text displayed in the view changes by adding the adornment to any reformatted lines
		/// </summary>
		/// <remarks><para>This event is raised whenever the rendered text displayed in the <see cref="MVSTE.ITextView">ITextView</see> changes.</para>
		/// <para>It is raised whenever the view does a layout (which happens when
		/// <see cref="MVSTE.ITextView.DisplayTextLineContainingBufferPosition">DisplayTextLineContainingBufferPosition</see> is called
		/// or in response to text or classification changes).</para>
		/// <para>It is also raised whenever the view scrolls horizontally or when its size changes.</para>
		/// <para>This function is called by the framework on Main Thread</para>
		/// </remarks>
		/// <param name="sender">The event sender.</param>
		/// <param name="e">The event arguments.</param>
		private void OnLayoutChanged(object sender, MVSTE.TextViewLayoutChangedEventArgs e)
		{
#if TRACE
			EnterFunction();
#endif
			StopTimer();

			MVST.ITextSnapshot text_snapshot = m_TextView.TextSnapshot;

#if TRACE_LAYOUT_CHANGED
			if (0 < e.NewOrReformattedLines.Count)
			{
				TraceMsg("--------");
				TraceMsg("NewOrReformattedLines.Count: " + e.NewOrReformattedLines.Count);
				TraceMsg("--------");
				foreach (MVSTF.ITextViewLine line in e.NewOrReformattedLines)
				{
					int line_number = text_snapshot.GetLineNumberFromPosition(line.Start);
					TraceMsg("Line " + line_number + " (" + line.Start.Position + "-" + line.End.Position + "): " + text_snapshot.GetText(line.Extent.Span));
				}
			}

			if (0 < e.TranslatedLines.Count)
			{
				TraceMsg("--------");
				TraceMsg("TranslatedLines.Count: " + e.TranslatedLines.Count);
				TraceMsg("--------");
				foreach (MVSTF.ITextViewLine line in e.TranslatedLines)
				{
					int line_number = text_snapshot.GetLineNumberFromPosition(line.Start);
					TraceMsg("Line " + line_number + " (" + line.Start.Position + "-" + line.End.Position + "): " + text_snapshot.GetText(line.Extent.Span));
				}
			}

			if (0 < e.NewOrReformattedSpans.Count)
			{
				TraceMsg("--------");
				TraceMsg("NewOrReformattedSpans.Count: " + e.NewOrReformattedSpans.Count);
				TraceMsg("--------");
				foreach (MVST.SnapshotSpan span in e.NewOrReformattedSpans)
				{
					int line_number = text_snapshot.GetLineNumberFromPosition(span.Start);
					TraceMsg("Span " + line_number + " (" + span.Start.Position + "-" + span.End.Position + "): " + text_snapshot.GetText(span.Span));
				}
			}

			if (0 < e.TranslatedSpans.Count)
			{
				TraceMsg("--------");
				TraceMsg("TranslatedSpans.Count: " + e.TranslatedSpans.Count);
				TraceMsg("--------");
				foreach (MVST.SnapshotSpan span in e.TranslatedSpans)
				{
					int line_number = text_snapshot.GetLineNumberFromPosition(span.Start);
					TraceMsg("Span " + line_number + " (" + span.Start.Position + "-" + span.End.Position + "): " + text_snapshot.GetText(span.Span));
				}
			}
#endif

			int first_line_number_to_add_adornment = int.MaxValue;
			int last_line_number_to_add_adornment  = int.MinValue;
			MVST.INormalizedTextChangeCollection changes = e.OldSnapshot.Version.Changes;
			if (null == changes || !changes.IncludesLineChanges || 0 == m_LineEntries.Count)
			{
				foreach (MVSTF.ITextViewLine line in e.NewOrReformattedLines)
				{
					MVST.SnapshotSpan snapshot_span = line.Extent;
					int line_number = text_snapshot.GetLineNumberFromPosition(line.Start);
					first_line_number_to_add_adornment = S.Math.Min(first_line_number_to_add_adornment, line_number);
					last_line_number_to_add_adornment  = S.Math.Max(last_line_number_to_add_adornment,  line_number);
					ProcessLine(snapshot_span, text_snapshot.GetText(snapshot_span.Span), text_snapshot.GetLineNumberFromPosition(line.Start));
				}
			}
			else
			{
#if DEBUG
				TRC_SD.Debug.Assert(e.OldSnapshot.Version.Next == e.NewSnapshot.Version);
#endif

				MVST.ITextSnapshot old_text_snapshot = e.OldSnapshot;

				// There could be more changes in a single line. Convert INormalizedTextChangeCollection changes to line based changes.
				var change_ranges = new ChangeRanges[changes.Count];
				int num_change_ranges = 0;
				foreach (MVST.ITextChange change in changes)
				{
					int old_line_number = old_text_snapshot.GetLineNumberFromPosition(change.OldPosition);
					int new_line_number = text_snapshot.GetLineNumberFromPosition(change.NewPosition);

					if (0 < num_change_ranges && old_line_number == change_ranges[num_change_ranges-1].m_OldLastLineNumber)
					{
						--num_change_ranges;
						change_ranges[num_change_ranges].m_LineCountDelta += change.LineCountDelta;
					}
					else
					{
						change_ranges[num_change_ranges].m_OldFirstLineNumber = old_line_number;
						change_ranges[num_change_ranges].m_NewFirstLineNumber = new_line_number;
						change_ranges[num_change_ranges].m_LineCountDelta = change.LineCountDelta;
					}

					for (;;)
					{
						++old_line_number;
						old_text_snapshot.GetLineNumberFromPosition(old_line_number);
						MVST.ITextSnapshotLine line = old_text_snapshot.GetLineFromLineNumber(old_line_number);
						if (change.OldEnd <= line.Start)
						{
							change_ranges[num_change_ranges].m_OldLastLineNumber = old_line_number - 1;
							break;
						}
					}

					for (;;)
					{
						++new_line_number;
						text_snapshot.GetLineNumberFromPosition(new_line_number);
						MVST.ITextSnapshotLine line = text_snapshot.GetLineFromLineNumber(new_line_number);
						if (change.NewEnd <= line.Start)
						{
							change_ranges[num_change_ranges].m_NewLastLineNumber = new_line_number - 1;
							break;
						}
					}

					++num_change_ranges;
				}

				var line_number_shifts = new LineNumberShift[m_LineEntries.Count];
				int change_ranges_index = 0;
				int num_line_number_shifts = 0;
				int line_count_delta = 0;
				foreach (SCG.KeyValuePair<int, LineEntry> pair in m_LineEntries)
				{
					int line_number = pair.Key;
					LineEntry line_entry = pair.Value;

					while (change_ranges_index < num_change_ranges && change_ranges[change_ranges_index].m_OldLastLineNumber < line_number)
					{
						line_count_delta += change_ranges[change_ranges_index].m_LineCountDelta;
						++change_ranges_index;
					}

					if (change_ranges_index < num_change_ranges && change_ranges[change_ranges_index].m_OldFirstLineNumber <= line_number && line_number <= change_ranges[change_ranges_index].m_OldLastLineNumber)
					{
						if (0 > change_ranges[change_ranges_index].m_LineCountDelta)
						{
							if (line_number + change_ranges[change_ranges_index].m_LineCountDelta < change_ranges[change_ranges_index].m_OldFirstLineNumber)
							{
								line_number_shifts[num_line_number_shifts].m_OldLineNumber = line_number;
								line_number_shifts[num_line_number_shifts].m_NewLineNumber = -1;
								line_number_shifts[num_line_number_shifts].m_LineEntry = line_entry;
								++num_line_number_shifts;
								continue;
							}
						}
					}

					if (0 != line_count_delta)
					{
						line_number_shifts[num_line_number_shifts].m_OldLineNumber = line_number;
						line_number_shifts[num_line_number_shifts].m_NewLineNumber = line_number + line_count_delta;
						line_number_shifts[num_line_number_shifts].m_LineEntry = line_entry;
						++num_line_number_shifts;
					}
				}

				for (int i = 0; i < num_line_number_shifts; ++i)
				{
					RemoveAdornment(line_number_shifts[i].m_OldLineNumber, line_number_shifts[i].m_LineEntry);
					m_LineEntries.Remove(line_number_shifts[i].m_OldLineNumber);
				}

				for (int i = 0; i < num_line_number_shifts; ++i)
				{
					if (0 < line_number_shifts[i].m_NewLineNumber)
					{
						m_LineEntries.Add(line_number_shifts[i].m_NewLineNumber, line_number_shifts[i].m_LineEntry);
						ShiftFigureLineNumber(line_number_shifts[i]);
					}
					else
					{
						RemoveFigure(line_number_shifts[i].m_OldLineNumber, line_number_shifts[i].m_LineEntry);
					}
				}

				for (int i = 0; i < num_change_ranges; ++i)
				{
					for (int line_number = change_ranges[i].m_NewFirstLineNumber; line_number <= change_ranges[i].m_NewLastLineNumber; ++line_number)
					{
						MVST.ITextSnapshotLine new_change_line = text_snapshot.GetLineFromLineNumber(line_number);
						ProcessLine(new_change_line.Extent, new_change_line.GetText(), line_number);
					}
				}
			}

			AddVisibleAdornments(first_line_number_to_add_adornment, last_line_number_to_add_adornment);

			if (0 < s_FigureLoadQueue.Count || s_CacheCanHaveUnreferencedEntries)
			{
				StartTimer();
			}
#if TRACE
			LeaveFunction();
#endif
		}

		/// <summary>Called when zoom is changed</summary>
		/// <remarks>This function is called by the framework on the MainThread</remarks>
		private void OnZoomLevelChanged(object sender, MVSTE.ZoomLevelChangedEventArgs e)
		{
#if TRACE
			EnterFunction();
#endif
			StopTimer();

			// Remove those lines from load queue, which are about to load with a different zoom level
			var figure_load_entries_to_remove = new SCG.List<FigureLoadQueueEntry>();
			foreach (FigureLoadQueueEntry figure_load_queue_entry in s_FigureLoadQueue)
			{
				FigureCacheEntry figure_cache_entry = figure_load_queue_entry.m_FigureCacheEntry;
				if (this == figure_load_queue_entry.m_Manager && figure_cache_entry.m_ZoomLevel != m_TextView.ZoomLevel)
				{
					figure_load_entries_to_remove.Add(figure_load_queue_entry);
				}
			}
			foreach (FigureLoadQueueEntry figure_load_queue_entry in figure_load_entries_to_remove)
			{
				s_FigureLoadQueue.Remove(figure_load_queue_entry);
			}

			foreach (SCG.KeyValuePair<int, LineEntry> pair in m_LineEntries)
			{
				LineEntry line_entry = pair.Value;
				int line_number = pair.Key;
				RemoveFigure(line_number, line_entry);
				var figure_cache_entry = new FigureCacheEntry(line_entry.m_FigureCacheEntry.m_FigurePath, m_TextView.ZoomLevel, line_entry.m_FigureCacheEntry.m_Inverted);
				if (s_FigureCache.TryGetValue(figure_cache_entry, out FigureCacheEntry stored_figure_cache_entry))
				{
					line_entry.m_FigureCacheEntry = stored_figure_cache_entry;
					FigureCacheData figure_cache_data = stored_figure_cache_entry.m_FigureCacheData;
					figure_cache_data.AddLineID(new LineID(this, line_number));
					AddFigure(line_number, line_entry, stored_figure_cache_entry.m_FigureCacheData);
				}
				else
				{
					figure_cache_entry.m_FigureCacheData = new FigureCacheData(new LineID(this, line_number));
					s_FigureCache.Add(figure_cache_entry);
					s_FigureLoadQueue.Add(new FigureLoadQueueEntry(this, figure_cache_entry, line_entry.m_ColorTheme));
					line_entry.m_FigureCacheEntry = figure_cache_entry;
				}
			}
			if (0 < s_FigureLoadQueue.Count || s_CacheCanHaveUnreferencedEntries)
			{
				StartTimer();
			}
#if TRACE
			LeaveFunction();
#endif
		}

#if TRACE
		internal static void TraceMsg(string message)
		{
			TRC_SD.Trace.Write("==== " + s_ThreadName.Value + " ");
			for (int i = 0; i < s_Indent.Value; ++i)
			{
				TRC_SD.Trace.Write(" ");
			}
			TRC_SD.Trace.WriteLine(message);
		}

		internal static void EnterFunction([TRC_SRCS.CallerMemberName] string function_name = "")
		{
			TraceMsg("Enter: " + function_name);
			s_Indent.Value += 2;
		}

		internal static void LeaveFunction([TRC_SRCS.CallerMemberName] string function_name = "")
		{
			s_Indent.Value -= 2;
			TraceMsg("Leave: " + function_name);
		}
#endif
	}

	/// <summary>
	/// Makes space for the adornment underneath the line
	/// </summary>
	internal class EmbedFigureLineTransformSource : MVSTF.ILineTransformSource
	{
		private readonly EmbedFigureManager m_Manager;

		internal EmbedFigureLineTransformSource(EmbedFigureManager manager)
		{
			m_Manager = manager;
		}

		/// <summary>
		/// Calculates the space needed at bottom of the line for the figure.
		/// </summary>
		/// <remarks>
		/// This function is called by the framework.
		/// </remarks>
		/// <param name="line">The line for which to calculate the line transform.</param>
		/// <param name="y_position">The y-coordinate of the line.</param>
		/// <param name="placement">The placement of the line with respect to y_position.</param>
		public MVSTF.LineTransform GetLineTransform(MVSTF.ITextViewLine line, double y_position, MVSTE.ViewRelativePosition placement)
		{
#if TRACE
			EmbedFigureManager.EnterFunction();
#endif
#if TRACE_LINE_TRANSFORM_LINE_NUMBERS
			EmbedFigureManager.TraceMsg("Transform line number: " + m_Manager.m_TextView.TextSnapshot.GetLineNumberFromPosition(line.Start.Position));
#endif
			MVSTF.LineTransform clt = line.LineTransform;
			MVSTF.LineTransform dlt = line.DefaultLineTransform;
			int line_number = line.Snapshot.GetLineNumberFromPosition(line.Start);

			double figure_height = 0.0;
			if (m_Manager.m_LineEntries.TryGetValue(line_number, out LineEntry line_entry))
			{
				if (null != line_entry.m_Figure)
				{
					figure_height = line_entry.m_Figure.m_Height;
				}
			}

#if TRACE
			EmbedFigureManager.LeaveFunction();
#endif

			return new MVSTF.LineTransform(clt.TopSpace, dlt.BottomSpace + figure_height, clt.VerticalScale);
		}
	}
}
