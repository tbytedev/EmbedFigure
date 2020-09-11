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

using MVST  = Microsoft.VisualStudio.Text;
using MVSTE = Microsoft.VisualStudio.Text.Editor;
using MVSTF = Microsoft.VisualStudio.Text.Formatting;
using MVSU  = Microsoft.VisualStudio.Utilities;
using SCMC  = System.ComponentModel.Composition;

namespace EmbedFigure
{
	/// <summary>
	/// Exports the <see cref="Microsoft.VisualStudio.Text.Editor.IWpfTextViewCreationListener"/> and <see cref="Microsoft.VisualStudio.Text.Editor.AdornmentLayerDefinition"/>
	/// Instantiates the <see cref="EmbedFigureManager"/> on the event of a <see cref="Microsoft.VisualStudio.Text.Editor.IWpfTextView"/>'s creation, if it wasn't already instantiated.
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
		/// Although it's not used directly, it's required when IWpfTextView.GetAdornmentLayer is called, and accessed via Export by the framework
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
		/// <param name="text_view">The view upon which the adornment should be placed</param>
		public void TextViewCreated(MVSTE.IWpfTextView text_view)
		{
			if (null == text_view || null == text_view.TextBuffer)
			{
				return;
			}

			// The adornment will listen to any event that changes the layout (text changes, scrolling, etc)
			EmbedFigureManager adornment = text_view.Properties.GetOrCreateSingletonProperty(() => new EmbedFigureManager(text_view, m_TextDocumentFactoryService));
		}
	}

	/// <summary>
	/// Exports the <see cref="Microsoft.VisualStudio.Text.Editor.IWpfTextViewCreationListener"/> and <see cref="Microsoft.VisualStudio.Text.Editor.AdornmentLayerDefinition"/>
	/// Instantiates the <see cref="EmbedFigureManager"/> on the event of a <see cref="EmbedFigureLineTransformSourceProvider.Create(MVSTE.IWpfTextView)"/>, if it wasn't already instantiated.
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
		/// Instantiates a <see cref="EmbedFigureManager"/> when the <see cref="MVSTF.ILineTransformSource"> is created, if it wasn't already instantiated.
		/// </summary>
		/// <param name="text_view">The view upon which the adornment should be placed</param>
		public MVSTF.ILineTransformSource Create(MVSTE.IWpfTextView text_view)
		{
			EmbedFigureManager adornment = text_view.Properties.GetOrCreateSingletonProperty(() => new EmbedFigureManager(text_view, m_TextDocumentFactoryService));
			return new EmbedFigureLineTransformSource(adornment);
		}
	}
}
