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

using SD   = System.Diagnostics;
using ST   = System.Threading;
using SRCS = System.Runtime.CompilerServices;

namespace EmbedFigure
{
#if TRACE
	internal static class Trace
	{
		private static readonly ST.ThreadLocal<string> s_ThreadName = new ST.ThreadLocal<string>(() => { return "Thread " + (10 > ST.Thread.CurrentThread.ManagedThreadId ? " " + ST.Thread.CurrentThread.ManagedThreadId : ST.Thread.CurrentThread.ManagedThreadId.ToString()); });
		private static readonly ST.ThreadLocal<int> s_Indent = new ST.ThreadLocal<int>(() => { return 0; });

		internal static void Message(string message)
		{
			SD.Trace.Write("==== " + s_ThreadName.Value + " ");
			for (int i = 0; i < s_Indent.Value; ++i)
			{
				SD.Trace.Write(" ");
			}
			SD.Trace.WriteLine(message);
		}

		internal static void EnterFunction([SRCS.CallerMemberName] string function_name = "")
		{
			Message("Enter: " + function_name);
			s_Indent.Value += 2;
		}

		internal static void LeaveFunction([SRCS.CallerMemberName] string function_name = "")
		{
			s_Indent.Value -= 2;
			Message("Leave: " + function_name);
		}
	}
#endif

#if DEBUG
	internal static class Debug
	{
		internal static void Assert(bool condition)
		{
			SD.Debug.Assert(condition);
		}
	}
#endif
}
