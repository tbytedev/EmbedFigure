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

using MVST  = Microsoft.VisualStudio.Text;
using MVSTT = Microsoft.VisualStudio.Text.Tagging;
using MVSTE = Microsoft.VisualStudio.Text.Editor;
using MVSTF = Microsoft.VisualStudio.Text.Formatting;
using MVSU  = Microsoft.VisualStudio.Utilities;
using SCMC  = System.ComponentModel.Composition;

namespace EmbedFigure
{
	/// <summary>
	/// Exports the <see cref="MVSTE.IWpfTextViewCreationListener"/> and <see cref="MVSTE.AdornmentLayerDefinition"/>, imports <see cref="MVST.ITextDocumentFactoryService"/>
	/// Instantiates the <see cref="EmbedFigureManager"/> on the event of a <see cref="MVSTE.IWpfTextView"/>'s creation, if it wasn't already instantiated.
	/// </summary>
	[MVSTE.TextViewRole(MVSTE.PredefinedTextViewRoles.Document)]
	[MVSU.ContentType("code")]
	[SCMC.Export(typeof(MVSTE.IWpfTextViewCreationListener))]
	internal class EmbedFigureTextViewCreationListener : MVSTE.IWpfTextViewCreationListener
	{
		// Disable "Field is never assigned to..." and "Field is never used" compiler's warnings. Justification: the field is used by MEF.
#pragma warning disable 169
		/// <summary>
		/// Defines the adornment layer for the adornment. This layer is ordered after the selection layer in the Z-order
		/// Although it's not used directly, it's required when IWpfTextView.GetAdornmentLayer is called, and accessed via Export by MEF
		/// </summary>
		[MVSU.Name("EmbedFigureAdornmentLayer")]
		[MVSU.Order(After = MVSTE.PredefinedAdornmentLayers.Selection, Before = MVSTE.PredefinedAdornmentLayers.Text)]
		[SCMC.Export]
		private readonly MVSTE.AdornmentLayerDefinition m_EditorAdornmentLayer;
#pragma warning restore 169

#pragma warning disable 649
		/// <summary>
		/// Service that creates, loads, and disposes text documents
		/// This variable is imported and never set by this code
		/// </summary>
		[SCMC.Import]
		private readonly MVST.ITextDocumentFactoryService m_TextDocumentFactoryService;
#pragma warning restore 649

		/// <summary>
		/// Called when a text view having matching roles is created over a text data model having a matching content type.
		/// Instantiates a <see cref="EmbedFigureManager"/> when the textView is created, if it wasn't already instantiated.
		/// </summary>
		/// <remarks>This function is called by the framework on Main Thread</remarks>
		/// <param name="text_view">The view upon which the adornment should be placed</param>
		public void TextViewCreated(MVSTE.IWpfTextView text_view)
		{
			if (null == text_view || null == text_view.TextBuffer)
			{
				return;
			}

			// The adornment will listen to any event that changes the layout (text changes, scrolling, etc)
			text_view.Properties.GetOrCreateSingletonProperty(() => new EmbedFigureManager(text_view, m_TextDocumentFactoryService));
		}
	}

	/// <summary>
	/// Exports the <see cref="MVSTF.ILineTransformSourceProvider"/> and imports <see cref="MVST.ITextDocumentFactoryService"/>
	/// Instantiates the <see cref="EmbedFigureManager"/> on the event of a <see cref="Create(MVSTE.IWpfTextView)"/>, if it wasn't already instantiated.
	/// </summary>
	[MVSTE.TextViewRole(MVSTE.PredefinedTextViewRoles.Interactive)]
	[MVSU.ContentType("code")]
	[SCMC.Export(typeof(MVSTF.ILineTransformSourceProvider))]
	internal class EmbedFigureLineTransformSourceProvider : MVSTF.ILineTransformSourceProvider
	{
#pragma warning disable 649
		/// <summary>
		/// Service that creates, loads, and disposes text documents
		/// This variable is imported and never set by this code
		/// </summary>
		[SCMC.Import]
		private readonly MVST.ITextDocumentFactoryService m_TextDocumentFactoryService;
#pragma warning restore 649

		/// <summary>
		/// Called when a text view having matching roles is created over a text data model having a matching content type.
		/// Instantiates a <see cref="EmbedFigureManager"/> when the <see cref="MVSTF.ILineTransformSource"/> is created, if it wasn't already instantiated.
		/// </summary>
		/// <remarks>This function is called by the framework on Main Thread</remarks>
		/// <param name="text_view">The view upon which the adornment should be placed</param>
		public MVSTF.ILineTransformSource Create(MVSTE.IWpfTextView text_view)
		{
			if (null == text_view || null == text_view.TextBuffer)
			{
				return null;
			}

			EmbedFigureManager manager = text_view.Properties.GetOrCreateSingletonProperty(() => new EmbedFigureManager(text_view, m_TextDocumentFactoryService));
			return new EmbedFigureLineTransformSource(manager);
		}
	}

	/// <summary>
	/// Exports the <see cref="MVSTT.ITaggerProvider"/>.
	/// </summary>
	[MVSTT.TagType(typeof(MVSTT.IErrorTag))]
	[MVSU.ContentType("code")]
	[MVSU.Name("EmbedFigure ErrorTagger")]
	[SCMC.Export(typeof(MVSTT.IViewTaggerProvider))]
	internal class ErrorTaggerProvider : MVSTT.IViewTaggerProvider
	{
		/// <summary>
		/// Called when the tagger for the specified <see cref="MVST.ITextBuffer"/> that has a matching content type should be created.
		/// </summary>
		/// <remarks>This function is called by the framework on Main Thread</remarks>
		/// <param name="text_view">The <see cref="MVSTE.ITextView"/> on which the tagger operates</param>
		/// <param name="text_buffer">The <see cref="MVST.ITextBuffer"/> on which the tagger operates</param>
		public MVSTT.ITagger<T> CreateTagger<T>(MVSTE.ITextView text_view, MVST.ITextBuffer text_buffer) where T : MVSTT.ITag
		{
			if (null == text_buffer)
			{
				return null;
			}

			return new EmbedFigureErrorTagger(text_view) as MVSTT.ITagger<T>;
		}
	}
}
