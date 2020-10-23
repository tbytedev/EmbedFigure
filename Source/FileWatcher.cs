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

using SCG = System.Collections.Generic;
using SIO = System.IO;

namespace EmbedFigure
{
	internal delegate void DirectoryWatcherEventHandler(WatchedDictionary sender, string filename);
	internal delegate void FileWatcherEventHandler(string dictionary_path, string filename);

	internal class WatchedDictionary
	{
		internal readonly SCG.Dictionary<string, int> m_Filenames = new SCG.Dictionary<string, int>();
		internal readonly SIO.FileSystemWatcher m_FileSystemWatcher = new SIO.FileSystemWatcher();

		internal event DirectoryWatcherEventHandler FileDeleted;
		internal event DirectoryWatcherEventHandler FileChanged;

		internal WatchedDictionary(string dictionary_path)
		{
			m_FileSystemWatcher.BeginInit();

			m_FileSystemWatcher.Path = dictionary_path;

			m_FileSystemWatcher.Changed += OnFileChanged;
			m_FileSystemWatcher.Created += OnFileChanged;
			m_FileSystemWatcher.Deleted += OnFileDeleted;
			m_FileSystemWatcher.Renamed += OnFileRenamed;

			// Begin watching
			m_FileSystemWatcher.EnableRaisingEvents = true;

			m_FileSystemWatcher.EndInit();
		}

		internal void AddFile(string filename)
		{
			if (m_Filenames.TryGetValue(filename, out int reference_count))
			{
				m_Filenames[filename] = reference_count + 1;
			}
			else
			{
				m_Filenames.Add(filename, 1);
			}
		}

		internal void RemoveFile(string filename)
		{
			if (m_Filenames.TryGetValue(filename, out int reference_count))
			{
				if (1 == reference_count)
				{
					m_Filenames.Remove(filename);
				}
				else
				{
					m_Filenames[filename] = reference_count - 1;
				}
			}
		}

		private void OnFileChanged(object sender, SIO.FileSystemEventArgs e)
		{
			string filename = e.Name.ToLowerInvariant();
			if (m_Filenames.ContainsKey(filename))
			{
				FileChanged(this, filename);
			}
		}

		private void OnFileDeleted(object sender, SIO.FileSystemEventArgs e)
		{
			string filename = e.Name.ToLowerInvariant();
			if (m_Filenames.ContainsKey(filename))
			{
				FileDeleted(this, filename);
			}
		}

		private void OnFileRenamed(object sender, SIO.RenamedEventArgs e)
		{
			OnFileDeleted(sender, new SIO.FileSystemEventArgs(SIO.WatcherChangeTypes.Deleted, e.OldFullPath, e.OldName));
			OnFileChanged(sender, new SIO.FileSystemEventArgs(SIO.WatcherChangeTypes.Created, e.FullPath, e.Name));
		}
	}

	internal class FileWatcher
	{
		internal readonly SCG.Dictionary<string, WatchedDictionary> m_WatchedDictionaries = new SCG.Dictionary<string, WatchedDictionary>();

		internal event FileWatcherEventHandler FileDeleted;
		internal event FileWatcherEventHandler FileChanged;

		internal void AddFile(string dictionary_path, string filename)
		{
			if (m_WatchedDictionaries.TryGetValue(dictionary_path, out WatchedDictionary watched_dictionary))
			{
				watched_dictionary.AddFile(filename);
			}
			else
			{
				watched_dictionary = new WatchedDictionary(dictionary_path);
				watched_dictionary.FileChanged += OnFileChanged;
				watched_dictionary.FileDeleted += OnFileDeleted;
				watched_dictionary.AddFile(filename);
				m_WatchedDictionaries.Add(dictionary_path, watched_dictionary);
			}
		}

		internal void RemoveFile(string dictionary_path, string filename)
		{
			if (m_WatchedDictionaries.TryGetValue(dictionary_path, out WatchedDictionary watched_dictionary))
			{
				watched_dictionary.RemoveFile(filename);
				if (0 == watched_dictionary.m_Filenames.Count)
				{
					watched_dictionary.m_FileSystemWatcher.Dispose();
					m_WatchedDictionaries.Remove(dictionary_path);
				}
			}
		}

		internal void OnFileChanged(WatchedDictionary sender, string filename)
		{
			FileChanged(sender.m_FileSystemWatcher.Path, filename);
		}

		internal void OnFileDeleted(WatchedDictionary sender, string filename)
		{
			FileDeleted(sender.m_FileSystemWatcher.Path, filename);
		}
	}
}
