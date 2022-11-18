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
//#define TRACE_LAYOUT_CHANGED
//#define TRACE_LAYOUT_SPAN_CHANGED
//#define TRACE_ADORNMENT_ADD_LINE_NUMBER
//#define TRACE_ADORNMENT_REMOVE_LINE_NUMBER
//#define TRACE_LINE_TRANSFORM_LINE_NUMBERS
#endif

#define HANDLE_FOCUS

using MVS   = Microsoft.VisualStudio;
using MVSS  = Microsoft.VisualStudio.Shell;
using MVSSI = Microsoft.VisualStudio.Shell.Interop;
using MVST  = Microsoft.VisualStudio.Text;
using MVSTE = Microsoft.VisualStudio.Text.Editor;
using MVSTF = Microsoft.VisualStudio.Text.Formatting;
using MVSTT = Microsoft.VisualStudio.Text.Tagging;
using S     = System;
using SCG   = System.Collections.Generic;
using SD    = System.Drawing;
using SDD2D = System.Drawing.Drawing2D;
using SDI   = System.Drawing.Imaging;
using SG    = System.Globalization;
using SIO   = System.IO;
using SRIS  = System.Runtime.InteropServices;
using ST    = System.Timers;
using STT   = System.Threading.Tasks;
using SW    = System.Windows;
using SWC   = System.Windows.Controls;
using SWM   = System.Windows.Media;
using SWMI  = System.Windows.Media.Imaging;

namespace EmbedFigure
{
	internal enum ColorTheme
	{
		Unspecified,
		Light,
		Dark
	}

	internal enum FigureSourceType
	{
		Unknown,
		ImageFile,
		SVGFile,
		TeXFile,
		TexString
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
	internal class Figure
	{
		internal readonly SWMI.BitmapSource m_BitmapSource;
		internal readonly double m_Height;

		internal Figure(SWMI.BitmapSource bitmap_source, double height)
		{
			m_BitmapSource = bitmap_source;
			m_Height = height;
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
		internal readonly string m_FigureSourceString;
		internal readonly FigureSourceType m_FigureSourceType;
		internal readonly double m_FigureScale;
		internal readonly bool m_Inverted;

		internal FigureCacheData m_FigureCacheData;

		internal FigureCacheEntry(string figure_source_string, FigureSourceType figure_source_type, double figure_scale, bool inverted)
		{
			m_FigureSourceString = figure_source_string;
			m_FigureSourceType = figure_source_type;
			m_FigureScale = figure_scale;
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
				&& m_FigureSourceString == entry.m_FigureSourceString
				&& m_FigureSourceType   == entry.m_FigureSourceType
				&& m_FigureScale        == entry.m_FigureScale
				&& m_Inverted           == entry.m_Inverted;
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
			hashCode = hashCode * -1521134295 + SCG.EqualityComparer<string>.Default.GetHashCode(m_FigureSourceString);
			hashCode = hashCode * -1521134295 + m_FigureSourceType.GetHashCode();
			hashCode = hashCode * -1521134295 + m_FigureScale.GetHashCode();
			hashCode = hashCode * -1521134295 + m_Inverted.GetHashCode();
			return hashCode;
		}

	}

	internal readonly struct FigureRenderQueueEntry
	{
		internal readonly EmbedFigureManager m_Manager;
		internal readonly FigureCacheEntry m_FigureCacheEntry;
		internal readonly ColorTheme m_ColorTheme;
		internal readonly double m_LineScale;

		internal FigureRenderQueueEntry(EmbedFigureManager manager, FigureCacheEntry figure_cache_entry, ColorTheme color_tone, double line_scale)
		{
			m_Manager = manager;
			m_FigureCacheEntry = figure_cache_entry;
			m_ColorTheme = color_tone;
			m_LineScale = line_scale;
		}
	}

	internal class LineEntry
	{
		/// <summary>
		/// Figure data for this line.
		/// </summary>
		/// <remarks>
		/// It's set right after the line is parsed, however the <see cref="Figure"/> reference inside this object can be null until the rendering has been finished.
		/// It's set to null if there's an error in this line.
		/// </remarks>
		internal FigureCacheEntry m_FigureCacheEntry;
		/// <summary>
		/// Stores data about the rendered figure
		/// </summary>
		/// <remarks>
		/// This is null, if the figure is not rendered yet. After the rendering is finished it's set to the proper value
		/// </remarks>
		internal Figure m_Figure;
		/// <summary>
		/// It there's an error in this line it contains the error message, otherwise it's null.
		/// </summary>
		internal string m_ErrorMessage;
		internal ColorTheme m_ColorTheme;
		internal double m_LineScale;
		internal int m_StartPositionIndex = -1;
		internal bool m_Added = false;

		internal LineEntry(ColorTheme color_theme, double line_scale)
		{
			m_ColorTheme = color_theme;
			m_LineScale = line_scale;
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
		ImageFile,
		Scale,
		SVGFile,
		TeXFile,
		TeXString,
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
		#region Constants

		private const double brightness_red   = 0.299;
		private const double brightness_green = 0.587;
		private const double brightness_blue  = 0.114;

		private const double max_scale = 10;
		private const double min_scale = 0.1;

		private const double initial_tex_scale = 20.0;
		private const string tex_font = "Arial";

		#endregion

		#region Static variables

		/// <summary>
		/// Stores rendered figures for each path in <see cref="SWMI.BitmapImage">BitmapImages</see>
		/// It's accessed only from Main thread
		/// </summary>
		internal static readonly SCG.HashSet<FigureCacheEntry> s_FigureCache = new SCG.HashSet<FigureCacheEntry>();

		internal static readonly SCG.List<EmbedFigureManager> s_Managers = new SCG.List<EmbedFigureManager>();

		/// <summary>
		/// Instruction to embed a figure to the source. This is prefixed by a # or @ prefix char
		/// </summary>
		private static readonly char[] s_InstructionCharArray = { 'E', 'm', 'b', 'e', 'd', 'F', 'i', 'g', 'u', 'r', 'e' };

		private static readonly ParameterToken[] s_ParameterDefinitions =
		{
			new ParameterToken("ColorTheme", ParameterType.ColorTheme),
			new ParameterToken("ImageFile",  ParameterType.ImageFile),
			new ParameterToken("Scale",      ParameterType.Scale),
			new ParameterToken("SVGFile",    ParameterType.SVGFile),
			new ParameterToken("TeXFile",    ParameterType.TeXFile),
			new ParameterToken("TeXString",  ParameterType.TeXString)
		};

		/// <summary>
		/// List of characters, that are not allowed in file names
		/// </summary>
		private static readonly char[] s_InvalidChars;

		/// <summary>
		/// This timer is fired after the user hasn't changed the text for 1500 ms and there are figures to render, or the cache can be cleaned up.
		/// </summary>
		/// <remarks>
		/// Do not load and render figures at once while user is still typing, rather wait some time to let things settle down a bit.
		/// Rendering is commenced when this timer is elapsed.
		/// </remarks>
		private static readonly ST.Timer s_Timer;

		/// <summary>
		/// Stores the figures to render and the lines to refresh
		/// It's accessed only from Main thread
		/// </summary>
		private static readonly SCG.HashSet<FigureRenderQueueEntry> s_FigureRenderQueue = new SCG.HashSet<FigureRenderQueueEntry>();

#if HANDLE_FOCUS
		private static readonly MVSSI.IVsShell s_VSShell;
		private static readonly VSEvents s_VSEvents;
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

		private static EmbedFigureOptions s_Options = EmbedFigureOptions.Instance;

		#endregion

		#region Static functions

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

			s_Options.OptionsChanged += OnOptionsChangedStatic;

			s_Timer = new ST.Timer(s_Options.UpdateDelay);
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
					InvalidateCache(figure_cache_data);
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
				SWM.Color color = solid_color_brush.Color;
				double brightness = GetPerceivedBrightness(color.ScR, color.ScG, color.ScB);
				if (0.5 < brightness)
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

		private static double GetPerceivedBrightness(double r, double g, double b)
		{
			return S.Math.Sqrt(brightness_red * r*r + brightness_green * g*g + brightness_blue * b*b);
		}

		private static void InvertBrightness(ref byte red, ref byte green, ref byte blue)
		{
			double sr = red   / 255.0;
			double sg = green / 255.0;
			double sb = blue  / 255.0;
			double source_brightness = GetPerceivedBrightness(sr, sg, sb);
			double target_brightness = 1.0 - source_brightness;
			if (0.5 < source_brightness)
			{
				double scale = 255.0 * target_brightness / source_brightness;
				red   = (byte)(scale * sr);
				green = (byte)(scale * sg);
				blue  = (byte)(scale * sb);
				return;
			}
			else
			{
				double max_component = S.Math.Max(S.Math.Max(sr, sg), sb);
				double mr;
				double mg;
				double mb;
				double middle_brightness;

				if (0.0 == max_component)
				{
					red = green = blue = 255;
					return;
				}

				if (1.0 > max_component)
				{
					double scale = 1 / max_component;
					mr = scale * sr;
					mg = scale * sg;
					mb = scale * sb;
					middle_brightness = GetPerceivedBrightness(mr, mg, mb);
				}
				else
				{
					mr = sr;
					mg = sg;
					mb = sb;
					middle_brightness = source_brightness;
				}

				if (middle_brightness < target_brightness)
				{
					double a = brightness_red * (1.0 - mr) * (1.0 - mr) + brightness_green * (1.0 - mg) * (1.0 - mg) + brightness_blue * (1.0 - mb) * (1.0 - mb);
					double b = 2.0 * brightness_red * mr * (1.0 - mr) + 2.0 * brightness_green * mg * (1.0 - mg) + 2.0 * brightness_blue * mb * (1.0 - mb);
					double c = brightness_red * mr * mr + brightness_green * mg * mg + brightness_blue * mb * mb - target_brightness * target_brightness;
					double d = b * b - 4 * a * c;
					double t = (-b + S.Math.Sqrt(d)) / (2 * a);
					double tr = (1 - t) * mr + t;
					double tg = (1 - t) * mg + t;
					double tb = (1 - t) * mb + t;
					red   = (byte)(255.0 * tr);
					green = (byte)(255.0 * tg);
					blue  = (byte)(255.0 * tb);
				}
				else
				{
					double scale = 255.0 * target_brightness / source_brightness;
					red   = (byte)(scale * sr);
					green = (byte)(scale * sg);
					blue  = (byte)(scale * sb);
				}
			}
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

		private static void RenderFigureTask(object context)
		{
#if TRACE_FUNCTIONS
			Trace.EnterFunction();
#endif
			var figure_cache_entry = (FigureCacheEntry)context;

			// Disposable objects
			SD.Bitmap bitmap = null;
			SD.Bitmap bitmap_temp = null;
			SD.Graphics graphics = null;
			SDI.ImageAttributes image_attributes = null;

			byte[] pixels = null;
			int bitmap_width = 0;
			int bitmap_height = 0;
			int bitmap_stride = 0;
			double image_height = 0.0;
			try
			{
				if (FigureSourceType.ImageFile == figure_cache_entry.m_FigureSourceType)
				{
					bitmap = new SD.Bitmap(figure_cache_entry.m_FigureSourceString);
					image_height = bitmap.Height;

					if (1.0 != figure_cache_entry.m_FigureScale)
					{
						bitmap_temp = bitmap;

						int scaled_width = S.Convert.ToInt32(bitmap_temp.Width * figure_cache_entry.m_FigureScale);
						int scaled_height = S.Convert.ToInt32(bitmap_temp.Height * figure_cache_entry.m_FigureScale);

						// Set up bitmap again for scaling
						bitmap = new SD.Bitmap(scaled_width, scaled_height);
						bitmap.SetResolution(bitmap_temp.HorizontalResolution, bitmap_temp.VerticalResolution);

						// Set up graphics
						graphics = SD.Graphics.FromImage(bitmap);
						graphics.InterpolationMode = SDD2D.InterpolationMode.High;
						graphics.CompositingQuality = SDD2D.CompositingQuality.HighQuality;
						graphics.SmoothingMode = SDD2D.SmoothingMode.AntiAlias;
						graphics.PixelOffsetMode = SDD2D.PixelOffsetMode.HighQuality;

						image_attributes = new SDI.ImageAttributes();
						image_attributes.SetWrapMode(SDD2D.WrapMode.TileFlipXY);
						var scaled_rectangle = new SD.Rectangle(0, 0, scaled_width, scaled_height);
						graphics.DrawImage(bitmap_temp, scaled_rectangle);
					}

					// Retrieve byte[] array that contains raw pixel data in BGRA format
					SDI.BitmapData bitmap_data = bitmap.LockBits(new SD.Rectangle(0, 0, bitmap.Width, bitmap.Height), SDI.ImageLockMode.ReadOnly, SDI.PixelFormat.Format32bppArgb);

					bitmap_width = bitmap_data.Width;
					bitmap_height = bitmap_data.Height;
					bitmap_stride = S.Math.Abs(bitmap_data.Stride);
					int bitmap_size = bitmap_stride * bitmap_height;
					pixels = new byte[bitmap_size];
					SRIS.Marshal.Copy(bitmap_data.Scan0, pixels, 0, bitmap_size);
					bitmap.UnlockBits(bitmap_data);
				}
				else if (FigureSourceType.SVGFile == figure_cache_entry.m_FigureSourceType)
				{
					Svg.SvgDocument svg_doc = Svg.SvgDocument.Open(figure_cache_entry.m_FigureSourceString);
					bitmap = svg_doc.Draw();
					image_height = bitmap.Height;

					// SVG library doesn't provide a way to retrieve the dimensions of the image before rendering.
					// But we can render the image and get the dimensions of the Bitmap.
					if (1.0 != figure_cache_entry.m_FigureScale)
					{
						// If zoom level is not 100%, render SVG file again in the resolution that matches the zoom level.
						int scaled_width = S.Convert.ToInt32(bitmap.Width * figure_cache_entry.m_FigureScale);
						int scaled_height = S.Convert.ToInt32(bitmap.Height * figure_cache_entry.m_FigureScale);
						bitmap.Dispose();   // Dispose bitmap before assigning new value
						bitmap = svg_doc.Draw(scaled_width, scaled_height);
					}

					// Retrieve byte[] array that contains raw pixel data in BGRA format
					SDI.BitmapData bitmap_data = bitmap.LockBits(new SD.Rectangle(0, 0, bitmap.Width, bitmap.Height), SDI.ImageLockMode.ReadOnly, SDI.PixelFormat.Format32bppArgb);

					bitmap_width = bitmap_data.Width;
					bitmap_height = bitmap_data.Height;
					bitmap_stride = S.Math.Abs(bitmap_data.Stride);
					int bitmap_size = bitmap_stride * bitmap_height;
					pixels = new byte[bitmap_size];
					SRIS.Marshal.Copy(bitmap_data.Scan0, pixels, 0, bitmap_size);
					bitmap.UnlockBits(bitmap_data);
				}
				else
				{
					string tex_file_content = null;
					if (FigureSourceType.TexString == figure_cache_entry.m_FigureSourceType)
					{
						tex_file_content = figure_cache_entry.m_FigureSourceString;
					}
					else if (FigureSourceType.TeXFile == figure_cache_entry.m_FigureSourceType)
					{
						// Read TeX file content
						tex_file_content = SIO.File.ReadAllText(figure_cache_entry.m_FigureSourceString);
					}

					// Parse TeX file
					var parser = new WpfMath.TexFormulaParser();
					WpfMath.TexFormula formula = parser.Parse(tex_file_content);

					MVSS.ThreadHelper.JoinableTaskFactory.Run(async delegate
					{
						await MVSS.ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

						// UI related objects (System.Windows.Media.Imaging.BitmapSource) can be created and used only on Main thread
						double tex_scale = initial_tex_scale * figure_cache_entry.m_FigureScale;
						WpfMath.TexRenderer renderer = formula.GetRenderer(WpfMath.TexStyle.Display, tex_scale, tex_font);
						SWMI.BitmapSource bitmap_source = renderer.RenderToBitmap(0.0, 0.0);

						bitmap_width = bitmap_source.PixelWidth;
						bitmap_height = bitmap_source.PixelHeight;
						bitmap_stride = bitmap_width * bitmap_source.Format.BitsPerPixel / 8;
						image_height = bitmap_source.Height / figure_cache_entry.m_FigureScale;
						int bitmap_size = bitmap_stride * bitmap_height;
						pixels = new byte[bitmap_size];
						bitmap_source.CopyPixels(pixels, bitmap_stride, 0);
					});
				}

				// Invert figure if needed
				if (figure_cache_entry.m_Inverted)
				{
					for (int y = 0; y < bitmap_height; ++y)
					{
						int i = bitmap_stride * y;
						for (int x = 0; x < bitmap_width; ++x, i += 4)
						{
							InvertBrightness(ref pixels[i+2], ref pixels[i+1], ref pixels[i]);
						}
					}
				}

#if TRACE_FUNCTIONS
				Trace.Message("Switch to Main RenderFigureTask");
#endif
				// UI related objects (System.Windows.Media.Imaging.BitmapSource) can be created and used only on Main thread
				MVSS.ThreadHelper.JoinableTaskFactory.Run(async delegate
				{
					await MVSS.ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

#if TRACE_FUNCTIONS
					Trace.Message("Switched to Main RenderFigureTask");
#endif
					FigureCacheData figure_cache_data = figure_cache_entry.m_FigureCacheData;
					if (0 == figure_cache_data.m_LineIDs.Count)
					{
						// This figure has been removed after the render task was started.
						return;
					}

					SWMI.BitmapSource bitmap_source = SWMI.BitmapSource.Create(bitmap_width, bitmap_height, 96, 96, SWM.PixelFormats.Bgra32, null, pixels, bitmap_stride);
					figure_cache_data.m_Figure = new Figure(bitmap_source, image_height);

					foreach (LineID line_id in figure_cache_data.m_LineIDs)
					{
						EmbedFigureManager manager = line_id.m_Manager;
						int line_number = line_id.m_LineNumber;
						LineEntry line_entry = manager.m_LineEntries[line_number];

						manager.AddFigure(line_number, line_entry, figure_cache_data);
					}
#if TRACE_FUNCTIONS
					Trace.Message("Switch from Main RenderFigureTask");
#endif
				});
			}
			// TODO: Error handling
			catch
			{
			}
			finally
			{
#if TRACE_FUNCTIONS
				Trace.Message("Switched from Main RenderFigureTask");
#endif
				bitmap?.Dispose();
				bitmap_temp?.Dispose();
				graphics?.Dispose();
				image_attributes?.Dispose();
			}

#if TRACE_FUNCTIONS
			Trace.LeaveFunction();
#endif
		}

		private static void ProcessLineRenderQueue()
		{
			// Create a list of lines for each figure. It's possible that more than one lines are waiting for the same figure to be rendered.
			var figure_render_queue = new SCG.Dictionary<FigureCacheEntry, SCG.HashSet<LineID>>();
			foreach (FigureRenderQueueEntry figure_render_queue_entry in s_FigureRenderQueue)
			{
				FigureCacheEntry figure_cache_entry = figure_render_queue_entry.m_FigureCacheEntry;
				FigureCacheData figure_cache_data = figure_cache_entry.m_FigureCacheData;

				// Check if there are still lines referring to this cache entry
				if (0 == figure_cache_data.m_LineIDs.Count)
				{
					continue;
				}

				// Check if figure source is a string
				if (FigureSourceType.TexString == figure_cache_entry.m_FigureSourceType)
				{
					// Check if this figure is already rendered
					if (null != figure_cache_data.m_Figure)
					{
						continue;
					}
				}
				else
				{
					// Check if update is necessary
					if (figure_cache_data.m_UpdateID == s_UpdateID)
					{
						continue;
					}

					figure_cache_data.m_UpdateID = s_UpdateID;

					var file_info = new SIO.FileInfo(figure_cache_entry.m_FigureSourceString);

					if (!file_info.Exists)
					{
						// This figure has been deleted.
						InvalidateCache(figure_cache_data);
						continue;
					}
					if (file_info.LastWriteTimeUtc == figure_cache_data.m_LastWriteTimeUtc)
					{
						continue;
					}
					InvalidateCache(figure_cache_data);
					figure_cache_data.m_LastWriteTimeUtc = file_info.LastWriteTimeUtc;
				}

				var task = new STT.Task(RenderFigureTask, figure_render_queue_entry.m_FigureCacheEntry);
				task.Start();
				// Although System.Threading.Tasks.Task is IDisposable, it's not necessary to call its Dispose() function.
				// https://devblogs.microsoft.com/pfxteam/do-i-need-to-dispose-of-tasks/
			}

			s_FigureRenderQueue.Clear();

		}

		private static void InvalidateCache(FigureCacheData figure_cache_data)
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
				figure_cache_data.m_Figure = null;
				figure_cache_data.m_LastWriteTimeUtc = S.DateTime.MinValue;
			}
		}

		private static void OnOptionsChangedStatic(object sender, S.EventArgs e)
		{
			if (s_Options.m_PrevUpdateDelay == s_Options.UpdateDelay)
			{
				return;
			}

			s_Timer.Interval = s_Options.UpdateDelay;
		}

		/// <summary>
		/// This timer is fired after the user hasn't changed the text for 1500 ms and there are figures to render.
		/// </summary>
		/// <remarks>This function is called by the framework on a Worker Thread</remarks>
		private static void TimerElapsed(int timer_start_id)
		{
#if TRACE_FUNCTIONS
			Trace.EnterFunction();
			Trace.Message("Switch to Main TimerElapsed");
#endif
			MVSS.ThreadHelper.JoinableTaskFactory.Run(async delegate
			{
				await MVSS.ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
				if (s_TimerStartID != timer_start_id)
				{
					return;
				}

#if TRACE_FUNCTIONS
				Trace.Message("Switched to Main TimerElapsed");
#endif
				ProcessLineRenderQueue();
				CacheCleanup();
#if TRACE_FUNCTIONS
				Trace.Message("Switch from Main TimerElapsed");
#endif
			});
#if TRACE_FUNCTIONS
			Trace.Message("Switched from Main TimerElapsed");
#endif
#if TRACE_FUNCTIONS
			Trace.LeaveFunction();
#endif
		}

		#endregion

		#region Member variables

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
		/// Stores line heights for line numbers. If there's no height stored for a line it means the height is 0.
		/// </summary>
		internal SCG.Dictionary<int, double> m_LastReturnedLineHeights = new SCG.Dictionary<int, double>();

		/// <summary>
		/// Current color theme
		/// </summary>
		private ColorTheme m_ColorTheme = ColorTheme.Unspecified;

		/// <summary>
		/// Indicates wether OnLayoutChanged should return immediately
		/// </summary>
		private bool SkipOnLayoutChanged = false;

		#endregion

		#region Member functions

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

			s_Options.OptionsChanged          += OnOptionsChanged;
		}

		private void AddAdornment(MVST.SnapshotSpan line_snapshot_span, int line_number, LineEntry line_entry)
		{
#if TRACE_FUNCTIONS
			Trace.EnterFunction();
#endif
			SWM.Geometry geometry = null;
			if (Align.Indentation == s_Options.Alignment)
			{
				string line_string = line_snapshot_span.GetText();
				int i = 0;
				while (char.IsWhiteSpace(line_string[i]) && i < line_snapshot_span.Length)
				{
					++i;
				}
				MVST.SnapshotSpan indented_snapshot_span = new MVST.SnapshotSpan(line_snapshot_span.Start + i, line_snapshot_span.End);
				if (0 < indented_snapshot_span.Length)
				{
					geometry = m_TextView.TextViewLines.GetMarkerGeometry(indented_snapshot_span);
				}
				else
				{
					geometry = m_TextView.TextViewLines.GetMarkerGeometry(line_snapshot_span);
				}
			}
			else
			{
				geometry = m_TextView.TextViewLines.GetMarkerGeometry(line_snapshot_span);
			}
			if (null != geometry)
			{
				var image = new SWC.Image
				{
					Source = line_entry.m_Figure.m_BitmapSource,
				};

				if (Align.Center == s_Options.Alignment)
				{
					SWC.Canvas.SetLeft(image, 0.5 * (m_TextView.ViewportWidth - line_entry.m_Figure.m_BitmapSource.Width));
				}
				else if (Align.Right == s_Options.Alignment)
				{
					SWC.Canvas.SetLeft(image, m_TextView.ViewportWidth - line_entry.m_Figure.m_BitmapSource.Width);
				}
				else
				{
					SWC.Canvas.SetLeft(image, geometry.Bounds.Left);
				}
				SWC.Canvas.SetTop(image, geometry.Bounds.Bottom);

				line_entry.m_Added = true;

				m_AdornmentLayer.AddAdornment(MVSTE.AdornmentPositioningBehavior.TextRelative, line_snapshot_span, line_number, image, OnAdornmentRemoved);

#if TRACE_ADORNMENT_ADD_LINE_NUMBER
				Trace.Message("Adornment add line number: " + line_number);
#endif
			}
#if TRACE_FUNCTIONS
			Trace.LeaveFunction();
#endif
		}

		private void AddVisibleAdornments(int first_line_number_to_add_adornment, int last_line_number_to_add_adornment)
		{
			if (m_TextView.TextViewLines.FormattedSpan.Start == m_TextView.TextViewLines.FormattedSpan.End)
			{
				// This file is empty, just return
				return;
			}

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
				MVSTE.TextViewExtensions.QueuePostLayoutAction(m_TextView, () => { UpdateLineHeight(line_number); AddAdornment(text_snapshot_line.Extent, line_number, line_entry); });
			}
		}

		private void AddFigure(int line_number, LineEntry line_entry, FigureCacheData figure_cache_data)
		{
			line_entry.m_Figure = figure_cache_data.m_Figure;
			if (null == line_entry.m_Figure)
			{
				return;
			}

			MVSTF.IWpfTextViewLine text_view_line = UpdateLineHeight(line_number);
			if (null != text_view_line)
			{
				AddAdornment(text_view_line.Extent, line_number, line_entry);
			}
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
		/// Search for #EmbedFigure instruction, and if found parse its parameters.
		/// </summary>
		/// <param name="line">Line to add the adornments</param>
		private string ParseLine(string line_string, out int out_start_position_index, out string out_figure_source_string, out FigureSourceType out_figure_source_type, out ColorTheme out_color_theme, out double out_line_scale)
		{
			out_figure_source_string = null;
			out_figure_source_type = FigureSourceType.Unknown;
			out_color_theme = ColorTheme.Unspecified;
			out_line_scale = 1.0;
			out_start_position_index = -1;
			int line_length = line_string.Length;

			int index_in_line;
			for (index_in_line = 0; index_in_line < line_length; ++index_in_line)
			{
			_again:
				if (s_Options.PrefixChar != line_string[index_in_line])
				{
					continue;
				}

				// s_Options.PrefixChar matters only if the previous character was not a letter or digit
				if (0 != index_in_line && char.IsLetterOrDigit(line_string[index_in_line - 1]))
				{
					continue;
				}

				int start_position_index = index_in_line;
				++index_in_line;
				// Compare instruction
				for (int index_in_instruction = 0; index_in_instruction < s_InstructionCharArray.Length; ++index_in_instruction, ++index_in_line)
				{
					if (index_in_line == line_length)
					{
						return null;
					}
					if (s_InstructionCharArray[index_in_instruction] != line_string[index_in_line])
					{
						goto _again;	// continue without incrementing the loop variable

					}
				}

				if (index_in_line == line_length)
				{
					out_start_position_index = start_position_index;
					return "No parameters specified.";
				}

				// Instruction must be followed by a whitespace
				if (!char.IsWhiteSpace(line_string[index_in_line]))
				{
					goto _again;	// continue without incrementing the loop variable
				}

				out_start_position_index = start_position_index;
				// Skip spaces between instruction and first parameter
				do
				{
					++index_in_line;
					if (index_in_line == line_length)
					{
						return "No parameters specified.";
					}
				}
				while (char.IsWhiteSpace(line_string[index_in_line]));

				SCG.List<ParameterToken> tokens = TokenizeParameters(line_string.Substring(index_in_line));

				ParameterType parameter_name = ParameterType.UnknownParameter;

				bool figure_source_specified = false;
				bool figure_scale_specified = false;

				foreach (ParameterToken token in tokens)
				{
					if (token.m_Type == ParameterType.ParameterValue)
					{
						switch (parameter_name)
						{
							case ParameterType.ImageFile:
							case ParameterType.SVGFile:
							case ParameterType.TeXFile:
								if (!figure_source_specified)
								{
									figure_source_specified = true;
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

									string local_path = token.m_RawText;
									// TODO: Check if file system is case sensitive.
									// For now we're assuming that it's a windows file system, so it's case insensitive.
									// Convert filename to lowercase to ignore character case.
									local_path = local_path.ToLowerInvariant();
									local_path = local_path.Replace(SIO.Path.AltDirectorySeparatorChar, SIO.Path.DirectorySeparatorChar);

									// TODO: Support for full path and other rooted path types
									out_figure_source_string = SIO.Path.GetDirectoryName(m_TextDocument.FilePath) + SIO.Path.DirectorySeparatorChar + local_path;

									if (ParameterType.ImageFile == parameter_name)
									{
										out_figure_source_type = FigureSourceType.ImageFile;
									}
									else if (ParameterType.SVGFile == parameter_name)
									{
										out_figure_source_type = FigureSourceType.SVGFile;
									}
									else
									{
										out_figure_source_type = FigureSourceType.TeXFile;
									}
								}
							_invalid_local_path:
								break;
							case ParameterType.TeXString:
								if (!figure_source_specified)
								{
									figure_source_specified = true;
									out_figure_source_string = token.m_RawText;
									out_figure_source_type = FigureSourceType.TexString;
								}
								break;
							case ParameterType.ColorTheme:
								if (ColorTheme.Unspecified == out_color_theme)
								{
									if ("dark" == token.m_RawText)
									{
										out_color_theme = ColorTheme.Dark;
									}
									else if ("light" == token.m_RawText)
									{
										out_color_theme = ColorTheme.Light;
									}
									else
									{
										return "Invalid ColorTheme : \"" + token.m_RawText + "\"";
									}
								}
								break;
							case ParameterType.Scale:
								if (!figure_scale_specified)
								{
									figure_scale_specified = true;
									int length = token.m_RawText.Length;
									try
									{
										if ('%' == token.m_RawText[length - 1])
										{
											out_line_scale = 0.01 * double.Parse(token.m_RawText.Substring(0, length - 1), SG.NumberStyles.Float);
										}
										else
										{
											out_line_scale = double.Parse(token.m_RawText, SG.NumberStyles.Float);
										}

										if (out_line_scale < min_scale)
										{
											return "Scale is less than the minimum scale (" + min_scale + ")";
										}
										else if (max_scale < out_line_scale)
										{
											return "Scale is greater than the maximum scale (" + max_scale + ")";
										}
									}
									catch
									{
										return "Invalid Scale : \"" + token.m_RawText + "\"";
									}
								}
								break;
						}
						parameter_name = ParameterType.UnknownParameter;
					}
					else
					{
						if (ParameterType.UnknownParameter == token.m_Type)
						{
							return "Unknown parameter : \"" + token.m_RawText + "\"";
						}
						parameter_name = token.m_Type;
					}
				}
				break;
			}
			return null;
		}

		/// <summary>
		/// Parse line. Search for #EmbedFigure instruction, and register figure changes.
		/// </summary>
		/// <param name="line">Line to add the adornments</param>
		private void ProcessLine(string line_text, int line_number)
		{
			string error_message = ParseLine(line_text, out int start_position_index, out string figure_source_string, out FigureSourceType figure_source_type, out ColorTheme color_theme, out double line_scale);

			var line_id = new LineID(this, line_number);

			if (null != error_message)
			{
				// The line is erroneous
				if (m_LineEntries.TryGetValue(line_number, out LineEntry line_entry))
				{
					if (null == line_entry.m_ErrorMessage)
					{
						// The line was correct before. We don't want to display anything if there's an error in the line, so the figure is not valid any more.
						RemoveFigure(line_number, line_entry);
					}
				}
				else
				{
					// Create new line entry for the error message
					line_entry = new LineEntry(color_theme, line_scale);
					m_LineEntries.Add(line_number, line_entry);
				}
				line_entry.m_ErrorMessage = error_message;
				line_entry.m_StartPositionIndex = start_position_index;
			}
			else if (null == figure_source_string)
			{
				// Currently there's no figure or error message specified in this line
				if (m_LineEntries.TryGetValue(line_number, out LineEntry line_entry))
				{
					// But there was a figure or an error message in this line previously. Remove line entry.
					RemoveFigure(line_number, line_entry);
					m_LineEntries.Remove(line_number);
				}
			}
			else
			{
				// There's a figure in this line
				double zoom_level = m_TextView.ZoomLevel / 100.0;
				double figure_scale = line_scale * zoom_level;
				bool inverted = ColorTheme.Unspecified != color_theme && color_theme != m_ColorTheme;

				if (m_LineEntries.TryGetValue(line_number, out LineEntry line_entry))
				{
					// There's a figure in this line now but there was already a figure in this line
					line_entry.m_ColorTheme = color_theme;
					line_entry.m_LineScale = line_scale;

					line_entry.m_ErrorMessage = null;
					line_entry.m_StartPositionIndex = -1;

					if (null                 == line_entry.m_FigureCacheEntry ||
						figure_source_string != line_entry.m_FigureCacheEntry.m_FigureSourceString ||
						figure_source_type   != line_entry.m_FigureCacheEntry.m_FigureSourceType ||
						inverted             != line_entry.m_FigureCacheEntry.m_Inverted ||
						figure_scale         != line_entry.m_FigureCacheEntry.m_FigureScale)
					{
						// The current and the previous figures are different
						RemoveFigure(line_number, line_entry);

						FigureCacheEntry figure_cache_entry = new FigureCacheEntry(figure_source_string, figure_source_type, figure_scale, inverted);

						if (s_FigureCache.TryGetValue(figure_cache_entry, out FigureCacheEntry stored_figure_cache_entry))
						{
							// This figure is already rendered, so just use it.
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
							s_FigureRenderQueue.Add(new FigureRenderQueueEntry(this, figure_cache_entry, color_theme, line_scale));
						}
					}
				}
				else
				{
					// There's a figure in this line and there was no figure in this line previously
					FigureCacheEntry figure_cache_entry = new FigureCacheEntry(figure_source_string, figure_source_type, figure_scale, inverted);
					line_entry = new LineEntry(color_theme, line_scale);
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
						s_FigureRenderQueue.Add(new FigureRenderQueueEntry(this, figure_cache_entry, color_theme, line_scale));
					}
					m_LineEntries.Add(line_number, line_entry);
				}
			}
		}

		/// <summary>
		/// Remove adornment from the UI
		/// </summary>
		/// <param name="line_number">Line number from where the adornment should be removed</param>
		/// <param name="line_entry">Corresponding line entry</param>
		private void RemoveAdornment(int line_number, LineEntry line_entry)
		{
			if (line_entry.m_Added)
			{
				m_AdornmentLayer.RemoveAdornmentsByTag(line_number);
			}
		}

		/// <summary>
		/// Remove adornment from the UI and unlink cache data from the line entry
		/// </summary>
		/// <param name="line_number">Line number from where the adornment should be removed</param>
		/// <param name="line_entry">Corresponding line entry</param>
		private void RemoveFigure(int line_number, LineEntry line_entry)
		{
			RemoveAdornment(line_number, line_entry);
			if (null != line_entry.m_FigureCacheEntry)
			{
				FigureCacheData figure_cache_data = line_entry.m_FigureCacheEntry.m_FigureCacheData;
				if (figure_cache_data.RemoveLineID(new LineID(this, line_number)))
				{
					s_CacheCanHaveUnreferencedEntries = true;
				}
				line_entry.m_FigureCacheEntry = null;
			}
			line_entry.m_Figure = null;
			MVSTE.TextViewExtensions.QueuePostLayoutAction(m_TextView, () => { UpdateLineHeight(line_number); });
		}

		/// <summary>
		/// Updates the height of the line.
		/// </summary>
		/// <remarks>
		/// <para>Makes enough space under the line for the adornment. It makes the framework call GetLineTransform which calculates the requires space.</para>
		/// <para>However there are some side effects. It also removes the adornment from the line and calls OnLayoutChanged.</para>
		/// </remarks>
		/// <param name="line_number">Line number to be updated</param>
		private MVSTF.IWpfTextViewLine UpdateLineHeight(int line_number)
		{
			// Check if the line height has been changed
			double target_line_height = 0.0;
			if (m_LineEntries.TryGetValue(line_number, out LineEntry line_entry))
			{
				if (null != line_entry.m_Figure)
				{
					target_line_height = line_entry.m_Figure.m_Height * line_entry.m_LineScale;
				}
			}

			m_LastReturnedLineHeights.TryGetValue(line_number, out double last_line_height);
			if (target_line_height == last_line_height)
			{
				// The line height hasn't been changed since it was set.
				return null;
			}

			// Get the first line from m_TextView.TextViewLines, this is not the same as m_TextView.TextViewLines.FirstVisibleLine.
			// It's possible that m_TextView.TextViewLines[0] is hidden and it's before m_TextView.TextViewLines.FirstVisibleLine
			MVSTF.IWpfTextViewLine first_view_line = m_TextView.TextViewLines[0];
			int first_view_line_number = m_TextView.TextSnapshot.GetLineNumberFromPosition(first_view_line.Start);
			int view_line_number = line_number - first_view_line_number;

			// Skip lines that has no corresponding line in m_TextView.TextViewLines
			if (0 > view_line_number)
			{
				return null;
			}
			if (m_TextView.TextViewLines.Count <= view_line_number)
			{
				return null;
			}

			MVSTF.IWpfTextViewLine text_view_line = m_TextView.TextViewLines[view_line_number];

			SkipOnLayoutChanged = true;
			// Force calling of GetLineTransform for this line. This makes space under the line to render the adornment.
			// Side effect: it also removes the adornment from the line and calls OnLayoutChanged
			m_TextView.DisplayTextLineContainingBufferPosition(text_view_line.Start, text_view_line.Top - m_TextView.ViewportTop, MVSTE.ViewRelativePosition.Top);
			SkipOnLayoutChanged = false;

			return text_view_line;
		}

		/// <summary>
		/// Called by the system after the adornment has been removed.
		/// </summary>
		/// <remarks>
		/// <para>The system automatically removes the adornments from those lines which are actually being edited.</para>
		/// </remarks>
		/// <param name="line_number">Line number from where the adornment should be removed</param>
		/// <param name="line_entry">Corresponding line entry</param>
		private void OnAdornmentRemoved(object tag, SW.UIElement element)
		{
#if TRACE_FUNCTIONS
			Trace.EnterFunction();
#endif
			int line_number = (int)tag;
			m_LineEntries[line_number].m_Added = false;
#if TRACE_ADORNMENT_REMOVE_LINE_NUMBER
			Trace.Message("Adornment remove line number: " + line_number);
#endif
#if TRACE_FUNCTIONS
			Trace.LeaveFunction();
#endif
		}

		private void OnBackgroundChanged(object sender, MVSTE.BackgroundBrushChangedEventArgs e)
		{
			ColorTheme color_theme = GetColorThemeFromBrush(e.NewBackgroundBrush);
			if (color_theme == m_ColorTheme)
			{
				return;
			}
#if TRACE_FUNCTIONS
			Trace.EnterFunction();
#endif
			m_ColorTheme = color_theme;

			StopTimer();

			// Remove those lines from render queue, which are about to render with a different invert value
			var figure_render_entries_to_remove = new SCG.List<FigureRenderQueueEntry>();
			foreach (FigureRenderQueueEntry figure_render_queue_entry in s_FigureRenderQueue)
			{
				FigureCacheEntry figure_cache_entry = figure_render_queue_entry.m_FigureCacheEntry;
				ColorTheme render_entry_color_theme = figure_render_queue_entry.m_ColorTheme;
				bool inverted = ColorTheme.Unspecified != render_entry_color_theme && render_entry_color_theme != m_ColorTheme;
				if (this == figure_render_queue_entry.m_Manager && figure_cache_entry.m_Inverted != inverted)
				{
					figure_render_entries_to_remove.Add(figure_render_queue_entry);
				}
			}
			foreach (FigureRenderQueueEntry figure_render_queue_entry in figure_render_entries_to_remove)
			{
				s_FigureRenderQueue.Remove(figure_render_queue_entry);
			}

			foreach (SCG.KeyValuePair<int, LineEntry> pair in m_LineEntries)
			{
				LineEntry line_entry = pair.Value;

				// Skip this line if it's only an error message and there's no figure.
				if (null != line_entry.m_ErrorMessage)
				{
					continue;
				}

				// If color theme is unspecified, leave the figure as it is
				if (ColorTheme.Unspecified == line_entry.m_ColorTheme)
				{
					continue;
				}

				bool inverted = line_entry.m_ColorTheme != m_ColorTheme;
				var figure_cache_entry = new FigureCacheEntry(line_entry.m_FigureCacheEntry.m_FigureSourceString,
				                                              line_entry.m_FigureCacheEntry.m_FigureSourceType,
				                                              line_entry.m_FigureCacheEntry.m_FigureScale,
				                                              inverted);

				int line_number = pair.Key;
				RemoveFigure(line_number, line_entry);

				if (s_FigureCache.TryGetValue(figure_cache_entry, out FigureCacheEntry stored_figure_cache_entry))
				{
					line_entry.m_FigureCacheEntry = stored_figure_cache_entry;
					FigureCacheData figure_cache_data = stored_figure_cache_entry.m_FigureCacheData;
					figure_cache_data.AddLineID(new LineID(this, line_number));
					line_entry.m_Figure = figure_cache_data.m_Figure;
				}
				else
				{
					line_entry.m_FigureCacheEntry = figure_cache_entry;
					figure_cache_entry.m_FigureCacheData = new FigureCacheData(new LineID(this, line_number));
					s_FigureCache.Add(figure_cache_entry);
					s_FigureRenderQueue.Add(new FigureRenderQueueEntry(this, figure_cache_entry, line_entry.m_ColorTheme, line_entry.m_LineScale));
				}
			}

			// We can process s_FigureRenderQueue this early, because background color changes only occasionally.
			// It's unlikely that it changes again before the previous rendering is finished.
			ProcessLineRenderQueue();
			AddVisibleAdornments(0, int.MaxValue);

			if (s_CacheCanHaveUnreferencedEntries)
			{
				StartTimer();
			}
#if TRACE_FUNCTIONS
			Trace.LeaveFunction();
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
#if TRACE_FUNCTIONS
			Trace.EnterFunction();
#endif
			StopTimer();

			s_FigureRenderQueue.Clear();

			foreach (SCG.KeyValuePair<int, LineEntry> pair in m_LineEntries)
			{
				LineEntry line_entry = pair.Value;
				if (null != line_entry.m_FigureCacheEntry)
				{
					s_FigureRenderQueue.Add(new FigureRenderQueueEntry(this, line_entry.m_FigureCacheEntry, line_entry.m_ColorTheme, line_entry.m_LineScale));
				}
			}

			ProcessLineRenderQueue();
#if TRACE_FUNCTIONS
			Trace.LeaveFunction();
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
			if (SkipOnLayoutChanged)
			{
				return;
			}

#if TRACE_FUNCTIONS
			Trace.EnterFunction();
#endif
			StopTimer();

			MVST.ITextSnapshot text_snapshot = m_TextView.TextSnapshot;

#if TRACE_LAYOUT_CHANGED
			if (0 < e.NewOrReformattedLines.Count)
			{
				Trace.Message("--------");
				Trace.Message("NewOrReformattedLines.Count: " + e.NewOrReformattedLines.Count);
				Trace.Message("--------");
				foreach (MVSTF.ITextViewLine line in e.NewOrReformattedLines)
				{
					int line_number = text_snapshot.GetLineNumberFromPosition(line.Start);
					Trace.Message("Line " + line_number + " (" + line.Start.Position + "-" + line.End.Position + "): " + text_snapshot.GetText(line.Extent.Span));
				}
			}

			if (0 < e.TranslatedLines.Count)
			{
				Trace.Message("--------");
				Trace.Message("TranslatedLines.Count: " + e.TranslatedLines.Count);
				Trace.Message("--------");
				foreach (MVSTF.ITextViewLine line in e.TranslatedLines)
				{
					int line_number = text_snapshot.GetLineNumberFromPosition(line.Start);
					Trace.Message("Line " + line_number + " (" + line.Start.Position + "-" + line.End.Position + "): " + text_snapshot.GetText(line.Extent.Span));
				}
			}

#if TRACE_LAYOUT_SPAN_CHANGED
			if (0 < e.NewOrReformattedSpans.Count)
			{
				Trace.Message("--------");
				Trace.Message("NewOrReformattedSpans.Count: " + e.NewOrReformattedSpans.Count);
				Trace.Message("--------");
				foreach (MVST.SnapshotSpan span in e.NewOrReformattedSpans)
				{
					int line_number = text_snapshot.GetLineNumberFromPosition(span.Start);
					Trace.Message("Span " + line_number + " (" + span.Start.Position + "-" + span.End.Position + "): " + text_snapshot.GetText(span.Span));
				}
			}

			if (0 < e.TranslatedSpans.Count)
			{
				Trace.Message("--------");
				Trace.Message("TranslatedSpans.Count: " + e.TranslatedSpans.Count);
				Trace.Message("--------");
				foreach (MVST.SnapshotSpan span in e.TranslatedSpans)
				{
					int line_number = text_snapshot.GetLineNumberFromPosition(span.Start);
					Trace.Message("Span " + line_number + " (" + span.Start.Position + "-" + span.End.Position + "): " + text_snapshot.GetText(span.Span));
				}
			}
#endif
#endif
			int first_line_number_to_add_adornment = int.MaxValue;
			int last_line_number_to_add_adornment  = int.MinValue;
			MVST.INormalizedTextChangeCollection changes = e.OldSnapshot.Version.Changes;
			if (null == changes || !changes.IncludesLineChanges || 0 == m_LineEntries.Count)
			{
				// Lines were not moved, so there was no line insertion or deletion
				foreach (MVSTF.ITextViewLine line in e.NewOrReformattedLines)
				{
					MVST.SnapshotSpan snapshot_span = line.Extent;
					int line_number = text_snapshot.GetLineNumberFromPosition(line.Start);
					first_line_number_to_add_adornment = S.Math.Min(first_line_number_to_add_adornment, line_number);
					last_line_number_to_add_adornment  = S.Math.Max(last_line_number_to_add_adornment,  line_number);
					ProcessLine(text_snapshot.GetText(snapshot_span.Span), text_snapshot.GetLineNumberFromPosition(line.Start));
				}
			}
			else
			{
				// There were line insertion or deletion, track line movements
#if DEBUG
				Debug.Assert(e.OldSnapshot.Version.Next == e.NewSnapshot.Version);
#endif
				first_line_number_to_add_adornment = 0;
				last_line_number_to_add_adornment = int.MaxValue;

				MVST.ITextSnapshot old_text_snapshot = e.OldSnapshot;

				SCG.IEnumerator<MVST.ITextChange> change_enumerator = changes.GetEnumerator();
				change_enumerator.MoveNext();
				MVST.ITextChange change = change_enumerator.Current;

				int line_count_delta = 0;

				var line_number_shifts = new LineNumberShift[m_LineEntries.Count];
				int num_line_number_shifts = 0;

				foreach (SCG.KeyValuePair<int, LineEntry> pair in m_LineEntries)
				{
					int line_number = pair.Key;
					LineEntry line_entry = pair.Value;
					MVST.ITextSnapshotLine line = old_text_snapshot.GetLineFromLineNumber(line_number);

					// Iterate through those changes that are before this line entity
					while (null != change && change.OldEnd <= line.Start)
					{
						line_count_delta += change.LineCountDelta;
						if (change_enumerator.MoveNext())
						{
							change = change_enumerator.Current;
						}
						else
						{
							change = null;
						}
					}

					if (null == change || line.EndIncludingLineBreak <= change.OldPosition)
					{
						// The current change is after this line entity. It means that it shouldn't be taken account yet.
						// Register the line number shift for the current line entity if the accumulated line delta is not 0,
						// and carry on to the next line entity.
						if (0 != line_count_delta)
						{
							line_number_shifts[num_line_number_shifts].m_OldLineNumber = line_number;
							line_number_shifts[num_line_number_shifts].m_NewLineNumber = line_number + line_count_delta;
							line_number_shifts[num_line_number_shifts].m_LineEntry = line_entry;
							++num_line_number_shifts;
						}
						continue;
					}

					// The current change and the current line entity overlap each other, so this line entity has been changed.
					// Just mark this line entity to be removed, It may be added again later in ProcessLine(), if the it's still valid.
					line_number_shifts[num_line_number_shifts].m_OldLineNumber = line_number;
					line_number_shifts[num_line_number_shifts].m_NewLineNumber = -1;	// -1 means that there's no new line number. This line entity must be removed.
					line_number_shifts[num_line_number_shifts].m_LineEntry = line_entry;
					++num_line_number_shifts;
				}

				change_enumerator.Dispose();

				// Remove line entities that are about to shift line and their adornments
				for (int i = 0; i < num_line_number_shifts; ++i)
				{
					RemoveAdornment(line_number_shifts[i].m_OldLineNumber, line_number_shifts[i].m_LineEntry);
					m_LineEntries.Remove(line_number_shifts[i].m_OldLineNumber);
				}

				for (int i = 0; i < num_line_number_shifts; ++i)
				{
					if (0 <= line_number_shifts[i].m_NewLineNumber)
					{
						m_LineEntries.Add(line_number_shifts[i].m_NewLineNumber, line_number_shifts[i].m_LineEntry);

						if (null != line_number_shifts[i].m_LineEntry.m_FigureCacheEntry)
						{
							FigureCacheData figure_cache_data = line_number_shifts[i].m_LineEntry.m_FigureCacheEntry.m_FigureCacheData;
							figure_cache_data.RemoveLineID(new LineID(this, line_number_shifts[i].m_OldLineNumber));
							figure_cache_data.AddLineID(new LineID(this, line_number_shifts[i].m_NewLineNumber));
						}
					}
					else
					{
						RemoveFigure(line_number_shifts[i].m_OldLineNumber, line_number_shifts[i].m_LineEntry);
					}
				}

				int last_line_number = 0;
				foreach (MVST.ITextChange curr_change in changes)
				{
					int line_number = text_snapshot.GetLineNumberFromPosition(curr_change.NewPosition);
					// There could be more changes in a single line.
					if (last_line_number != line_number)
					{
						MVST.ITextSnapshotLine line = text_snapshot.GetLineFromLineNumber(line_number);
						ProcessLine(line.GetText(), line_number);
						last_line_number = line_number;
					}

					for (;;)
					{
						++line_number;
						MVST.ITextSnapshotLine line = text_snapshot.GetLineFromLineNumber(line_number);
						if (curr_change.NewEnd <= line.Start)
						{
							break;
						}
						ProcessLine(line.GetText(), line_number);
						last_line_number = line_number;
					}
				}
			}

			AddVisibleAdornments(first_line_number_to_add_adornment, last_line_number_to_add_adornment);

			if (0 < s_FigureRenderQueue.Count || s_CacheCanHaveUnreferencedEntries)
			{
				StartTimer();
			}
#if TRACE_FUNCTIONS
			Trace.LeaveFunction();
#endif
		}

		private void OnOptionsChanged(object sender, S.EventArgs e)
		{
			if (s_Options.m_PrevPrefixChar == s_Options.PrefixChar)
			{
				return;
			}

			StopTimer();

			s_FigureRenderQueue.Clear();
			foreach (SCG.KeyValuePair<int, LineEntry> pair in m_LineEntries)
			{
				LineEntry line_entry = pair.Value;
				int line_number = pair.Key;
				RemoveFigure(line_number, line_entry);
			}
			m_LineEntries.Clear();

			if (m_TextView.TextViewLines.FormattedSpan.Start == m_TextView.TextViewLines.FormattedSpan.End)
			{
				// This file is empty, just return
				return;
			}

			int first_visible_line_number = m_TextView.TextSnapshot.GetLineNumberFromPosition(m_TextView.TextViewLines.FormattedSpan.Start);
			int last_visible_line_number = m_TextView.TextSnapshot.GetLineNumberFromPosition(m_TextView.TextViewLines.FormattedSpan.End - 1);
			for (int i = first_visible_line_number; i <= last_visible_line_number; ++i)
			{
				MVST.ITextSnapshotLine line = m_TextView.TextSnapshot.GetLineFromLineNumber(i);
				ProcessLine(line.GetText(), i);
			}

			AddVisibleAdornments(first_visible_line_number, last_visible_line_number);

			if (0 < s_FigureRenderQueue.Count || s_CacheCanHaveUnreferencedEntries)
			{
				StartTimer();
			}
		}

		/// <summary>Called when zoom is changed</summary>
		/// <remarks>This function is called by the framework on the MainThread</remarks>
		private void OnZoomLevelChanged(object sender, MVSTE.ZoomLevelChangedEventArgs e)
		{
#if TRACE_FUNCTIONS
			Trace.EnterFunction();
#endif
			StopTimer();

			double zoom_level = m_TextView.ZoomLevel / 100.0;
			// Remove those lines from render queue, which are about to be rendered with a different zoom level
			var figure_render_entries_to_remove = new SCG.List<FigureRenderQueueEntry>();
			foreach (FigureRenderQueueEntry figure_render_queue_entry in s_FigureRenderQueue)
			{
				FigureCacheEntry figure_cache_entry = figure_render_queue_entry.m_FigureCacheEntry;
				if (this == figure_render_queue_entry.m_Manager && figure_cache_entry.m_FigureScale != zoom_level * figure_render_queue_entry.m_LineScale)
				{
					figure_render_entries_to_remove.Add(figure_render_queue_entry);
				}
			}
			foreach (FigureRenderQueueEntry figure_render_queue_entry in figure_render_entries_to_remove)
			{
				s_FigureRenderQueue.Remove(figure_render_queue_entry);
			}

			// Update line entities
			foreach (SCG.KeyValuePair<int, LineEntry> pair in m_LineEntries)
			{
				LineEntry line_entry = pair.Value;

				// Skip this line if it's only an error message and there's no figure.
				if (null != line_entry.m_ErrorMessage)
				{
					continue;
				}

				var figure_cache_entry = new FigureCacheEntry(line_entry.m_FigureCacheEntry.m_FigureSourceString,
				                                              line_entry.m_FigureCacheEntry.m_FigureSourceType,
				                                              line_entry.m_LineScale * zoom_level,
				                                              line_entry.m_FigureCacheEntry.m_Inverted);

				int line_number = pair.Key;
				RemoveFigure(line_number, line_entry);

				if (s_FigureCache.TryGetValue(figure_cache_entry, out FigureCacheEntry stored_figure_cache_entry))
				{
					line_entry.m_FigureCacheEntry = stored_figure_cache_entry;
					FigureCacheData figure_cache_data = stored_figure_cache_entry.m_FigureCacheData;
					figure_cache_data.AddLineID(new LineID(this, line_number));
					line_entry.m_Figure = figure_cache_data.m_Figure;
				}
				else
				{
					figure_cache_entry.m_FigureCacheData = new FigureCacheData(new LineID(this, line_number));
					s_FigureCache.Add(figure_cache_entry);
					s_FigureRenderQueue.Add(new FigureRenderQueueEntry(this, figure_cache_entry, line_entry.m_ColorTheme, line_entry.m_LineScale));
					line_entry.m_FigureCacheEntry = figure_cache_entry;
				}
			}

			AddVisibleAdornments(0, int.MaxValue);

			if (0 < s_FigureRenderQueue.Count || s_CacheCanHaveUnreferencedEntries)
			{
				StartTimer();
			}
#if TRACE_FUNCTIONS
			Trace.LeaveFunction();
#endif
		}

		#endregion
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
#if TRACE_FUNCTIONS
			Trace.EnterFunction();
#endif
			MVSTF.LineTransform clt = line.LineTransform;
			MVSTF.LineTransform dlt = line.DefaultLineTransform;
			int line_number = line.Snapshot.GetLineNumberFromPosition(line.Start);

			double adornment_height = 0.0;
			if (m_Manager.m_LineEntries.TryGetValue(line_number, out LineEntry line_entry))
			{
				if (null != line_entry.m_Figure)
				{
					adornment_height = line_entry.m_Figure.m_Height * line_entry.m_LineScale;
				}
			}

			if (0.0 < adornment_height)
			{
				m_Manager.m_LastReturnedLineHeights[line_number] = adornment_height;
			}
			else
			{
				m_Manager.m_LastReturnedLineHeights.Remove(line_number);
			}

#if TRACE_LINE_TRANSFORM_LINE_NUMBERS
			if (clt.BottomSpace < dlt.BottomSpace + adornment_height)
			{
				Trace.Message("Transform line number: " + m_Manager.m_TextView.TextSnapshot.GetLineNumberFromPosition(line.Start.Position) + " height inc: " + dlt.BottomSpace + adornment_height);
			}
			else if (clt.BottomSpace > dlt.BottomSpace + adornment_height)
			{
				Trace.Message("Transform line number: " + m_Manager.m_TextView.TextSnapshot.GetLineNumberFromPosition(line.Start.Position) + " height dec: " + dlt.BottomSpace + adornment_height);
			}
			else
			{
				Trace.Message("Transform line number: " + m_Manager.m_TextView.TextSnapshot.GetLineNumberFromPosition(line.Start.Position) + " height equ: " + dlt.BottomSpace + adornment_height);
			}

#endif
#if TRACE_FUNCTIONS
			Trace.LeaveFunction();
#endif
			return new MVSTF.LineTransform(clt.TopSpace, dlt.BottomSpace + adornment_height, clt.VerticalScale);
		}
	}

	internal class EmbedFigureErrorTagger : MVSTT.ITagger<MVSTT.ErrorTag>
	{
		private EmbedFigureManager m_Manager;

#pragma warning disable 67
		/// <summary>
		/// We need this declaration to fulfill ITagger interface requirement.
		/// We should raise this event when tags are changed, but as there's no listener for this, we just leave this unused.
		/// </summary>
		public event S.EventHandler<MVST.SnapshotSpanEventArgs> TagsChanged;
#pragma warning restore 67

		internal EmbedFigureErrorTagger(MVSTE.ITextView text_view)
		{
			foreach (EmbedFigureManager manager in EmbedFigureManager.s_Managers)
			{
				if (manager.m_TextView == text_view)
				{
					m_Manager = manager;
					return;
				}
			}
		}

		public SCG.IEnumerable<MVSTT.ITagSpan<MVSTT.ErrorTag>> GetTags(MVST.NormalizedSnapshotSpanCollection spans)
		{
			var tag_list = new SCG.List<MVSTT.TagSpan<MVSTT.ErrorTag>>();

			MVSTE.IWpfTextView text_view = m_Manager.m_TextView;

			int first_visible_line_number = text_view.TextSnapshot.GetLineNumberFromPosition(text_view.TextViewLines.FormattedSpan.Start);
			int last_visible_line_number = text_view.TextSnapshot.GetLineNumberFromPosition(text_view.TextViewLines.FormattedSpan.End - 1);

			SCG.IEnumerator<MVST.SnapshotSpan> span_enumerator = spans.GetEnumerator();
			if (!span_enumerator.MoveNext())
			{
				return tag_list;
			}

			foreach (SCG.KeyValuePair<int, LineEntry> pair in m_Manager.m_LineEntries)
			{
				int line_number = pair.Key;
				LineEntry line_entry = pair.Value;

				if (line_number < first_visible_line_number)
				{
					// Skip those lines that precede the first visible line
					continue;
				}

				if (line_number > last_visible_line_number)
				{
					// Skip the rest of the lines, because they succeed the last visible line
					return tag_list;
				}

				if (null == line_entry.m_ErrorMessage)
				{
					// There's no error message in this line, so just skip
					continue;
				}

				MVST.SnapshotSpan line_span = text_view.TextSnapshot.GetLineFromLineNumber(line_number).Extent;
				MVST.SnapshotSpan line_error_span = new MVST.SnapshotSpan(line_span.Start.Add(line_entry.m_StartPositionIndex), line_span.End);

				while (span_enumerator.Current.End < line_error_span.Start)
				{
					// Skip those spans that precede the the current line
					if (!span_enumerator.MoveNext())
					{
						// Return if we reached the last span
						return tag_list;
					}
				}

				if (span_enumerator.Current.Start > line_error_span.End)
				{
					// Proceed to the next line if the current span succeed the current line
					continue;
				}

				MVST.SnapshotSpan? error_span = line_error_span.Intersection(span_enumerator.Current);
				if (error_span.HasValue)
				{
					var error_tag = new MVSTT.ErrorTag("EmbedFigure error type", "EmbedFigure | " + line_entry.m_ErrorMessage);
					MVSTT.TagSpan<MVSTT.ErrorTag> tag_span = new MVSTT.TagSpan<MVSTT.ErrorTag>(error_span.Value, error_tag);
					tag_list.Add(tag_span);
				}
			}
			return tag_list;
		}
	}
}
