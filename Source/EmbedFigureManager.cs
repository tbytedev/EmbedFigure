/*
 * EmbedFigure - Visual Studio extension for embedding math figures into source code
 * Copyright(C) 2020 Tamas Kezdi
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

#define TRACE

using MVST  = Microsoft.VisualStudio.Text;
using MVSTE = Microsoft.VisualStudio.Text.Editor;
using MVSTF = Microsoft.VisualStudio.Text.Formatting;
using SC    = System.Collections;
using SCG   = System.Collections.Generic;
using SD    = System.Drawing;
using SDI   = System.Drawing.Imaging;
using SIO   = System.IO;
using SW    = System.Windows;
using SWC   = System.Windows.Controls;
using SWM   = System.Windows.Media;
using SWMI  = System.Windows.Media.Imaging;

#if TRACE
using System.Diagnostics;
using System.Threading;
using System.Runtime.CompilerServices;
#endif

namespace EmbedFigure
{
	enum ColorTheme
	{
		Unspecified,
		Light,
		Dark
	}

	/// <summary>
	/// Contains the rendered figure, that is ready to add to the <see cref="Microsoft.VisualStudio.Text.Editor.IAdornmentLayer">IAdornmentLayer</see>
	/// </summary>
	internal class Figure
	{
		public readonly double m_ZoomLevel;
		public readonly ColorTheme m_ColorTheme;
		public readonly bool m_Inverted;

		private readonly string m_FigurePath;

		public double m_Height;
		public SWMI.BitmapImage m_BitmapImage;

		private int m_ReferenceCount = 1;

		public Figure(string figure_path, double zoom_level, ColorTheme color_tone, bool inverted)
		{
			m_FigurePath = figure_path;
			m_ZoomLevel = zoom_level;
			m_ColorTheme = color_tone;
			m_Inverted = inverted;
		}

		public void GenerateImage()
		{
			try
			{
				Svg.SvgDocument svg_doc = Svg.SvgDocument.Open(m_FigurePath);
				SD.Bitmap bitmap = svg_doc.Draw();
				m_Height = bitmap.Height;
				if (1.0 != m_ZoomLevel)
				{
					bitmap = svg_doc.Draw(System.Convert.ToInt32(bitmap.Width * m_ZoomLevel), System.Convert.ToInt32(bitmap.Height * m_ZoomLevel));
				}

				if (m_Inverted)
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
				// Save System.Drawing.Bitmap to a Create System.IO.MemoryStream
				var memory_stream = new SIO.MemoryStream();
				bitmap.Save(memory_stream, SDI.ImageFormat.Tiff);
				memory_stream.Seek(0, SIO.SeekOrigin.Begin);

				// UI related objects (System.Windows.Media.Imaging.BitmapImage) can be created and used only on Main thread
				Microsoft.VisualStudio.Shell.ThreadHelper.JoinableTaskFactory.Run(async delegate
				{
					await Microsoft.VisualStudio.Shell.ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

#if TRACE
					EmbedFigureManager.TraceMsg("Switch to Main GenerateImage");
#endif

					// Load back bitmap as System.Windows.Media.Imaging.BitmapImage from System.IO.MemoryStream
					m_BitmapImage = new SWMI.BitmapImage();
					m_BitmapImage.BeginInit();
					m_BitmapImage.StreamSource = memory_stream;
					m_BitmapImage.EndInit();

					EmbedFigureManager.s_FigureCache[new FigureCacheKey(m_FigurePath, m_Inverted)] = this;
				});
			}
			catch
			{
				m_BitmapImage = null;
				m_Height = 0.0;
			}
		}

		public void IncreaseReferenceCount()
		{
			++m_ReferenceCount;
		}

		public void DecreaseReferenceCount()
		{
			--m_ReferenceCount;
		}

		public int GetReferenceCount()
		{
			return m_ReferenceCount;
		}
	}

	internal readonly struct FigureCacheKey
	{
		public FigureCacheKey(string figure_path, bool inverted)
		{
			m_FigurePath = figure_path;
			m_Inverted = inverted;
		}
		public readonly string m_FigurePath;
		public readonly bool m_Inverted;
	}

	internal readonly struct FigureParams
	{
		public FigureParams(string figure_path, ColorTheme color_tone, bool inverted)
		{
			m_FigurePath = figure_path;
			m_ColorTheme = color_tone;
			m_Inverted = inverted;
		}
		public readonly string m_FigurePath;
		public readonly ColorTheme m_ColorTheme;
		public readonly bool m_Inverted;
	}

	internal class LineEntry
	{
		public LineEntry(string figure_path, bool inverted)
		{
			m_FigurePath = figure_path;
			m_Inverted = inverted;
			m_Figure = null;
			m_Added = false;
		}
		public string m_FigurePath;
		public Figure m_Figure;
		public bool m_Added;
		public bool m_Inverted;
	}

	internal readonly struct LineId
	{
		public LineId(EmbedFigureManager manager, int line_number)
		{
			m_Manager = manager;
			m_LineNumber = line_number;
		}
		public readonly EmbedFigureManager m_Manager;
		public readonly int m_LineNumber;
	}

	internal readonly struct LineUpdateInfo
	{
		public LineUpdateInfo(LineId line_id, Figure figure)
		{
			m_LineId = line_id;
			m_Figure = figure;
		}
		public readonly LineId m_LineId;
		public readonly Figure m_Figure;
	}

	enum ParameterType
	{
		UnknownParameter,
		ColorTheme,
		SVGFile,
		ParameterValue
	}

	internal readonly struct ParameterToken
	{
		public ParameterToken(string raw_text, ParameterType type)
		{
			m_RawText = raw_text;
			m_Type = type;
		}
		public readonly string m_RawText;
		public readonly ParameterType m_Type;
	}

	/// <summary>
	/// TextAdornment to place figures after #EmbedFigure instructions
	/// </summary>
	internal class EmbedFigureManager
	{
		/// <summary>
		/// Stores rendered figures for each path in <see cref="System.Windows.Media.Imaging.BitmapImage">BitmapImages</see>
		/// It's accessed only from Main thread
		/// </summary>
		internal static readonly SCG.Dictionary<FigureCacheKey, Figure> s_FigureCache = new SCG.Dictionary<FigureCacheKey, Figure>();

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
		/// This timer is fired after the user hasn't changed the text for 1500 ms.
		/// </summary>
		/// <remarks>
		/// Do not load and render figures at once while user is still typing, rather wait some time to let things settle down a bit.
		/// Load is commenced when this timer is elapsed.
		/// </remarks>
		private static readonly System.Timers.Timer s_LoadingTimer = new System.Timers.Timer(500);

		/// <summary>
		/// Stores the figures to load and the lines to refresh
		/// It's accessed from both Main and Timer thread
		/// </summary>
		private static readonly SCG.Dictionary<LineId, FigureParams> s_LineLoadQueue = new SCG.Dictionary<LineId, FigureParams>();

#if TRACE
		private static readonly ThreadLocal<string> s_ThreadName = new ThreadLocal<string>(() => { return "Thread " + (10 > Thread.CurrentThread.ManagedThreadId ? " " + Thread.CurrentThread.ManagedThreadId : Thread.CurrentThread.ManagedThreadId.ToString()); });
		private static readonly ThreadLocal<int> s_Indent = new ThreadLocal<int>(() => { return 0; });
#endif

		/// <summary>
		/// Indicates if the timer is active.
		/// </summary>
		/// <remarks>
		/// <see cref="OnTimerElapsed">OnTimerElapsed</see> may be called even after <see cref="s_LoadingTimer">s_LoadingTimer</see> has been stopped.
		/// It's set to true when the timer is started, and set to false when the timer is stopped. <see cref="OnTimerElapsed">OnTimerElapsed</see>
		/// checks this flag and if it's false, returns immediately.
		/// </remarks>
		private static bool s_LoadingTimerActive = false;

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
			System.Array.Sort(s_InvalidChars);

			s_LoadingTimer.AutoReset = false;
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

		private static void CleanUpCache()
		{
			var keys_to_remove = new SCG.List<FigureCacheKey>();
			foreach (SCG.KeyValuePair<FigureCacheKey, Figure> pair in s_FigureCache)
			{
				Figure figure = pair.Value;
				if (0 == figure.GetReferenceCount())
				{
					keys_to_remove.Add(pair.Key);
				}
			}

			foreach (FigureCacheKey key in keys_to_remove)
			{
				s_FigureCache.Remove(key);
			}
		}

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
		private readonly MVSTE.IWpfTextView m_TextView;

		/// <summary>
		/// Stores figure info for each line in this manager
		/// </summary>
		internal SCG.Dictionary<int, LineEntry> m_LineFigures = new SCG.Dictionary<int, LineEntry>();

		/// <summary>
		/// Current color theme
		/// </summary>
		private ColorTheme m_BackgroundTone = ColorTheme.Unspecified;

		/// <summary>
		/// Initializes a new instance of the <see cref="EmbedFigureManager"/> class.
		/// </summary>
		/// <param name="text_view">Text view to create the adornment for</param>
		internal EmbedFigureManager(MVSTE.IWpfTextView text_view, MVST.ITextDocumentFactoryService text_document_factory_service)
		{
			if (text_view == null)
			{
				throw new System.ArgumentNullException("TextView");
			}

			if (!text_document_factory_service.TryGetTextDocument(text_view.TextBuffer, out m_TextDocument))
			{
				// Document doesn't exist for textView.TextBuffer
				return;
			}

			m_AdornmentLayer = text_view.GetAdornmentLayer("EmbedFigureAdornmentLayer");

			m_TextView = text_view;

			// Detect background color brightness
			m_BackgroundTone = GetColorThemeFromBrush(m_TextView.Background);

			m_TextView.BackgroundBrushChanged += OnBackgroundChanged;
			m_TextView.LayoutChanged          += OnLayoutChanged;
			m_TextView.ZoomLevelChanged       += OnZoomLevelChanged;

			s_LoadingTimer.Elapsed            += OnTimerElapsed;
		}

		private void AddAdornment(MVSTF.ITextViewLine line, int line_number, LineEntry line_entry)
		{
#if TRACE
			EnterFunction();
#endif
			MVST.SnapshotSpan span = line.Extent;
			SWM.Geometry geometry = m_TextView.TextViewLines.GetMarkerGeometry(span);
			if (null != geometry)
			{
				var image = new SWC.Image
				{
					Source = line_entry.m_Figure.m_BitmapImage,
					Height = line_entry.m_Figure.m_Height
				};

				SWC.Canvas.SetLeft(image, geometry.Bounds.Left);
				SWC.Canvas.SetTop(image, geometry.Bounds.Bottom);
				m_AdornmentLayer.AddAdornment(MVSTE.AdornmentPositioningBehavior.TextRelative, span, line_number, image, OnAdornmentRemoved);
				line_entry.m_Added = true;
			}
#if TRACE
			LeaveFunction();
#endif
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

				string figure_filename = null;

				foreach (ParameterToken token in tokens)
				{
					if (token.m_Type == ParameterType.ParameterValue)
					{
						switch (parameter_name)
						{
							case ParameterType.SVGFile:
								if (null == figure_filename)
								{
									if (0 == token.m_RawText.Length)
									{
										break;
									}

									// Filename mustn't contain invalid characters
									foreach (char c in token.m_RawText)
									{
										if (0 <= System.Array.BinarySearch(s_InvalidChars, c))
										{
											goto _invalid_filename;
										}
									}

									char last_char = token.m_RawText[token.m_RawText.Length - 1];
									if (SIO.Path.DirectorySeparatorChar == last_char || SIO.Path.AltDirectorySeparatorChar == last_char)
									{
										// Last character is a directory separator. It can't be a filename, it's a directory.
										break;
									}

									figure_filename = token.m_RawText;
								}
							_invalid_filename:
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

				if (null == figure_filename)
				{
					return;
				}
				// TODO: Check if file system is case sensitive.
				// For now we're assuming that it's a windows file system, so it's case insensitive.
				// Convert filename to lowercase to ignore character case.
				figure_filename = figure_filename.ToLower();
				figure_filename = figure_filename.Replace(SIO.Path.AltDirectorySeparatorChar, SIO.Path.DirectorySeparatorChar);

				// TODO: Support for full path and other rooted path types
				out_figure_path = string.Concat(SIO.Path.GetDirectoryName(m_TextDocument.FilePath), SIO.Path.DirectorySeparatorChar, figure_filename);
				return;
			}
		}

		/// <summary>
		/// Parse line. Search for #EmbedFigure instruction, and register figure changes.
		/// </summary>
		/// <param name="line">Line to add the adornments</param>
		private void ProcessLine(MVSTF.ITextViewLine line)
		{

			MVST.ITextSnapshot text_snapshot = m_TextView.TextSnapshot;
			ParseLine(text_snapshot.GetText(line.Extent.Span), out string figure_path, out ColorTheme color_tone);

			//MVSTE.IWpfTextViewLineCollection text_view_lines = m_TextView.TextViewLines;
			int line_number = text_snapshot.GetLineNumberFromPosition(line.Start);
			var line_id = new LineId(this, line_number);

			lock ((s_LineLoadQueue as SC.ICollection).SyncRoot)
			{
				s_LineLoadQueue.Remove(line_id);
			}

			if (null == figure_path)
			{
				// There's no figure specified in this line currently
				if (m_LineFigures.TryGetValue(line_number, out LineEntry old_line_entry))
				{
					// But there was a figure in this line previously
					if (null != old_line_entry.m_Figure)
					{
						old_line_entry.m_Figure.DecreaseReferenceCount();
						old_line_entry.m_Figure = null;
					}
					m_LineFigures.Remove(line_number);
				}
			}
			else
			{
				bool inverted = ColorTheme.Unspecified != color_tone && color_tone != m_BackgroundTone;
				FigureCacheKey figure_cache_key = new FigureCacheKey(figure_path, inverted);

				if (m_LineFigures.TryGetValue(line_number, out LineEntry line_entry))
				{
					if (line_entry.m_FigurePath != figure_path || line_entry.m_Inverted != inverted || null == line_entry.m_Figure)
					{
						if (null != line_entry.m_FigurePath)
						{
							if (null != line_entry.m_Figure)
							{
								line_entry.m_Figure.DecreaseReferenceCount();
								line_entry.m_Figure = null;
							}
						}

						line_entry.m_FigurePath = figure_path;
						line_entry.m_Inverted = inverted;

						// Figure path has been changed in this line
						if (s_FigureCache.TryGetValue(figure_cache_key, out Figure figure) && figure.m_ZoomLevel == m_TextView.ZoomLevel)
						{
							figure.IncreaseReferenceCount();
							line_entry.m_Figure = figure;
							AddAdornment(line, line_number, line_entry);
						}
						else
						{
							lock ((s_LineLoadQueue as SC.ICollection).SyncRoot)
							{
								s_LineLoadQueue[line_id] = new FigureParams(figure_path, color_tone, inverted);
							}
						}
					}
					else if (!line_entry.m_Added)
					{
						AddAdornment(line, line_number, line_entry);
					}
				}
				else
				{
					line_entry = new LineEntry(figure_path, inverted);
					m_LineFigures[line_number] = line_entry;
					if (s_FigureCache.TryGetValue(figure_cache_key, out Figure figure) && figure.m_ZoomLevel == m_TextView.ZoomLevel)
					{
						line_entry.m_Figure = figure;
						AddAdornment(line, line_number, line_entry);
					}
					else
					{
						lock ((s_LineLoadQueue as SC.ICollection).SyncRoot)
						{
							s_LineLoadQueue[line_id] = new FigureParams(figure_path, color_tone, inverted);
						}
					}
				}
			}
		}

		private void OnAdornmentRemoved(object tag, SW.UIElement element)
		{
#if TRACE
			EnterFunction();
#endif
			int line_number = (int)tag;
			m_LineFigures[line_number].m_Added = false;
#if TRACE
			LeaveFunction();
#endif
		}

		private void OnBackgroundChanged(object sender, MVSTE.BackgroundBrushChangedEventArgs e)
		{
#if TRACE
			EnterFunction();
#endif
			s_LoadingTimerActive = false;
			s_LoadingTimer.Stop();

			ColorTheme background_tone = GetColorThemeFromBrush(e.NewBackgroundBrush);
			if (background_tone != m_BackgroundTone)
			{
				m_BackgroundTone = background_tone;
			}

			lock ((s_LineLoadQueue as SC.ICollection).SyncRoot)
			{
				foreach (SCG.KeyValuePair<int, LineEntry> pair in m_LineFigures)
				{
					LineEntry line_entry = pair.Value;
					int line_number = pair.Key;
					m_AdornmentLayer.RemoveAdornmentsByTag(line_number);
					if (null != line_entry.m_Figure)
					{
						ColorTheme color_tone = line_entry.m_Figure.m_ColorTheme;
						bool inverted = ColorTheme.Unspecified != color_tone && color_tone != m_BackgroundTone;
						s_LineLoadQueue[new LineId(this, line_number)] = new FigureParams(line_entry.m_FigurePath, color_tone, inverted);

						line_entry.m_Figure.DecreaseReferenceCount();
						line_entry.m_Figure = null;
					}
				}
			}
			s_LoadingTimerActive = true;
			s_LoadingTimer.Start();
#if TRACE
			LeaveFunction();
#endif
		}

		/// <summary>
		/// Handles whenever the text displayed in the view changes by adding the adornment to any reformatted lines
		/// </summary>
		/// <remarks><para>This event is raised whenever the rendered text displayed in the <see cref="Microsoft.VisualStudio.Text.Editor.ITextView">ITextView</see> changes.</para>
		/// <para>It is raised whenever the view does a layout (which happens when <see cref="Microsoft.VisualStudio.Text.Editor.ITextView.DisplayTextLineContainingBufferPosition">DisplayTextLineContainingBufferPosition</see> is called or in response to text or classification changes).</para>
		/// <para>It is also raised whenever the view scrolls horizontally or when its size changes.</para>
		/// </remarks>
		/// <param name="sender">The event sender.</param>
		/// <param name="e">The event arguments.</param>
		private void OnLayoutChanged(object sender, MVSTE.TextViewLayoutChangedEventArgs e)
		{
			if (0 == e.NewOrReformattedLines.Count)
			{
				return;
			}
#if TRACE
			EnterFunction();
#endif
			s_LoadingTimerActive = false;
			s_LoadingTimer.Stop();

			foreach (MVSTF.ITextViewLine line in e.NewOrReformattedLines)
			{
				ProcessLine(line);
			}

			lock ((s_LineLoadQueue as SC.ICollection).SyncRoot)
			{
				if (0 != s_LineLoadQueue.Count)
				{
					s_LoadingTimerActive = true;
					s_LoadingTimer.Start();
				}
			}
#if TRACE
			LeaveFunction();
#endif
		}

		private void OnTimerElapsed(object sender, System.Timers.ElapsedEventArgs e)
		{
			if (!s_LoadingTimerActive)
			{
				return;
			}

#if TRACE
			EnterFunction();
#endif
			var lines_to_update = new SCG.List<LineUpdateInfo>();

			// Create a list of lines for each figure. It's possible that more than one lines are waiting for the same figure to be loaded.
			var figure_load_queue = new SCG.Dictionary<FigureParams, SCG.List<LineId>>();
			lock ((s_LineLoadQueue as SC.ICollection).SyncRoot)
			{
				foreach (SCG.KeyValuePair<LineId, FigureParams> pair in s_LineLoadQueue)
				{
					FigureParams figure_params = pair.Value;
					LineId line_id = pair.Key;
					if (figure_load_queue.TryGetValue(figure_params, out SCG.List<LineId> line_list))
					{
						line_list.Add(line_id);
					}
					else
					{
						line_list = new SCG.List<LineId>
						{
							line_id
						};
						figure_load_queue.Add(figure_params, line_list);
					}
				}
			}

			// Iterate through figures, and try to load them. Create a list of those lines that can be refreshed after their figures have been loaded.
			// s_LineLoadQueue can't be locked, because Figure.GenerateImage switches back to Main thread, and if Main is currently waiting for s_LineLoadQueue, there will be a deadlock
			foreach (SCG.KeyValuePair<FigureParams, SCG.List<LineId>> pair in figure_load_queue)
			{
				FigureParams figure_params = pair.Key;
				string figure_path = figure_params.m_FigurePath;

				if (!SIO.File.Exists(figure_path))
				{
					continue;
				}

				Figure figure = new Figure(figure_path, m_TextView.ZoomLevel / 100.0, figure_params.m_ColorTheme, figure_params.m_Inverted);
				figure.GenerateImage();

				if (null != figure.m_BitmapImage)
				{
					SCG.List<LineId> line_list = pair.Value;
					foreach (LineId line_id in line_list)
					{
						lines_to_update.Add(new LineUpdateInfo(line_id, figure));
					}
				}
			}

			lock ((s_LineLoadQueue as SC.ICollection).SyncRoot)
			{
				// Remove those lines of which the figure has just been loaded
				foreach (LineUpdateInfo line_id_figure in lines_to_update)
				{
					s_LineLoadQueue.Remove(line_id_figure.m_LineId);
				}
			}

			// Switch to Main thread.
			// We access s_Figures only on Main thread, and updating IWpfTextView should also happen on Main thread
			Microsoft.VisualStudio.Shell.ThreadHelper.JoinableTaskFactory.Run(async delegate
			{
				await Microsoft.VisualStudio.Shell.ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

#if TRACE
				TraceMsg("Switch to Main OnTimerElapsed");
#endif

				CleanUpCache();

				// Refresh those lines of which the figure has just been loaded
				foreach (LineUpdateInfo line_id_figure in lines_to_update)
				{
					EmbedFigureManager manager = line_id_figure.m_LineId.m_Manager;
					int line_number = line_id_figure.m_LineId.m_LineNumber;
					MVSTE.IWpfTextView text_view = manager.m_TextView;

					// Get the first line from text_view.TextViewLines, this is not the same as text_view.TextViewLines.FirstVisibleLine.
					// It's possible that text_view.TextViewLines[0] is hidden and it's before text_view.TextViewLines.FirstVisibleLine
					MVSTF.IWpfTextViewLine first_view_line = text_view.TextViewLines[0];
					int first_view_line_number = text_view.TextSnapshot.GetLineNumberFromPosition(first_view_line.Start);
					int view_line_number = line_number - first_view_line_number;

					// Skip lines that has no corresponding line in text_view.TextViewLines
					if (0 > view_line_number)
					{
						continue;
					}
					if (manager.m_TextView.TextViewLines.Count <= view_line_number)
					{
						continue;
					}

					// Refresh line
					LineEntry line_entry = m_LineFigures[line_number];
					line_entry.m_Figure = line_id_figure.m_Figure;

					MVSTF.IWpfTextViewLine curr_view_line = manager.m_TextView.TextViewLines[view_line_number];
					AddAdornment(curr_view_line, line_number, line_entry);
					// Forces the call of GetLineTransform for this line
					manager.m_TextView.DisplayTextLineContainingBufferPosition(curr_view_line.Start, curr_view_line.Top - text_view.ViewportTop, MVSTE.ViewRelativePosition.Top);
				}

				if (0 != s_LineLoadQueue.Count)
				{
					s_LoadingTimerActive = true;
					s_LoadingTimer.Start();
				}
			});
#if TRACE
			LeaveFunction();
#endif
		}

		private void OnZoomLevelChanged(object sender, MVSTE.ZoomLevelChangedEventArgs e)
		{
#if TRACE
			EnterFunction();
#endif
			s_LoadingTimerActive = false;
			s_LoadingTimer.Stop();

			lock ((s_LineLoadQueue as SC.ICollection).SyncRoot)
			{
				foreach(SCG.KeyValuePair<int, LineEntry> pair in m_LineFigures)
				{
					LineEntry line_entry = pair.Value;
					int line_number = pair.Key;
					m_AdornmentLayer.RemoveAdornmentsByTag(line_number);
					if (null != line_entry.m_Figure)
					{
						s_LineLoadQueue[new LineId(this, line_number)] = new FigureParams(line_entry.m_FigurePath, line_entry.m_Figure.m_ColorTheme, line_entry.m_Figure.m_Inverted);

						line_entry.m_Figure.DecreaseReferenceCount();
						line_entry.m_Figure = null;
					}
				}
			}
			s_LoadingTimerActive = true;
			s_LoadingTimer.Start();
#if TRACE
			LeaveFunction();
#endif
		}

#if TRACE
		internal static void TraceMsg(string message)
		{
			Trace.Write("==== " + s_ThreadName.Value + " ");
			for (int i = 0; i < s_Indent.Value; ++i)
			{
				Trace.Write(" ");
			}
			Trace.WriteLine(message);
		}

		internal static void EnterFunction([CallerMemberName] string function_name = "")
		{
			TraceMsg("Enter: " + function_name);
			s_Indent.Value += 2;
		}
		internal static void LeaveFunction([CallerMemberName] string function_name = "")
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
		private readonly EmbedFigureManager m_Adornment;

		internal EmbedFigureLineTransformSource(EmbedFigureManager adornment)
		{
			m_Adornment = adornment;
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

			MVSTF.LineTransform clt = line.LineTransform;
			MVSTF.LineTransform dlt = line.DefaultLineTransform;
			int line_number = line.Snapshot.GetLineNumberFromPosition(line.Start);

			double figure_height = 0.0;
			if (m_Adornment.m_LineFigures.TryGetValue(line_number, out LineEntry line_entry))
			{
				if (null != line_entry.m_Figure && line_entry.m_Added)
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
