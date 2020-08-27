using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Task = System.Threading.Tasks.Task;
using System.Linq;

namespace SpacesTabifier {
	/// <summary>
	/// Command handler
	/// </summary>
	internal sealed class CommandReplaceSpaces {
		/// <summary>
		/// Command ID.
		/// </summary>
		public const int CommandId = 0x0100;

		/// <summary>
		/// Command menu group (command set GUID).
		/// </summary>
		public static readonly Guid CommandSet = new Guid("173db9f5-5808-4c90-9167-326995b8aa5a");

		/// <summary>
		/// VS Package that provides this command, not null.
		/// </summary>
		private readonly AsyncPackage package;

		/// <summary>
		/// Initializes a new instance of the <see cref="CommandReplaceSpaces"/> class.
		/// Adds our command handlers for menu (commands must exist in the command table file)
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		/// <param name="commandService">Command service to add command to, not null.</param>
		private CommandReplaceSpaces(AsyncPackage package, OleMenuCommandService commandService) {
			this.package = package ?? throw new ArgumentNullException(nameof(package));
			commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

			var menuCommandID = new CommandID(CommandSet, CommandId);
			var menuItem = new MenuCommand(this.Execute, menuCommandID);
			commandService.AddCommand(menuItem);
		}

		/// <summary>
		/// Gets the instance of the command.
		/// </summary>
		public static CommandReplaceSpaces Instance {
			get;
			private set;
		}

		/// <summary>
		/// Gets the service provider from the owner package.
		/// </summary>
		private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider {
			get {
				return this.package;
			}
		}

		/// <summary>
		/// Initializes the singleton instance of the command.
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		public static async Task InitializeAsync(AsyncPackage package) {
			// Switch to the main thread - the call to AddCommand in Command1's constructor requires
			// the UI thread.
			await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

			OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
			Instance = new CommandReplaceSpaces(package, commandService);
		}

		/// <summary>
		/// This function is the callback used to execute the command when the menu item is clicked.
		/// See the constructor to see how the menu item is associated with this function using
		/// OleMenuCommandService service and MenuCommand class.
		/// </summary>
		/// <param name="sender">Event sender.</param>
		/// <param name="e">Event args.</param>
		private void Execute(object sender, EventArgs e) {
			ThreadHelper.ThrowIfNotOnUIThread();

			var componentModel = (IComponentModel)Package.GetGlobalService(typeof(SComponentModel));
			var textManager = (IVsTextManager)Package.GetGlobalService(typeof(SVsTextManager));
			IVsTextView activeView = null;

			if (ErrorHandler.Failed(textManager.GetActiveView(1, null, out activeView))) {
				VsShellUtilities.ShowMessageBox(
					(IServiceProvider)this.ServiceProvider,
					"Can`t find opened document!",
					"Error",
					OLEMSGICON.OLEMSGICON_CRITICAL,
					OLEMSGBUTTON.OLEMSGBUTTON_OK,
					OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
				return;
			}

			var editorAdapter = componentModel.GetService<IVsEditorAdaptersFactoryService>();
			var textView = editorAdapter.GetWpfTextView(activeView);

			int tabSize = textView.Options.GetOptionValue(DefaultOptions.TabSizeOptionId);

			var snapshot = textView.TextSnapshot;
			var lines = GetLines(snapshot, textView).Where(x => x.GetText().Contains(" ")).ToArray();	// выдираем те, что с пробелами

			using (var edit = snapshot.TextBuffer.CreateEdit()) {
				foreach (var line in lines) {
					var newLine = Replacer(line.GetText(), tabSize);
					if (line.GetText() != newLine) {
						if (!edit.Replace(line.Start.Position, line.Length, newLine)) return;
					}
				}

				edit.Apply();
			}
		}

		IEnumerable<ITextSnapshotLine> GetLines(ITextSnapshot snapshot, IWpfTextView textView) {
			int start = snapshot.GetLineNumberFromPosition(textView.Selection.Start.Position);
			int end = snapshot.GetLineNumberFromPosition(textView.Selection.End.Position);

			if (start == end) {
				start = 0;
				end = snapshot.LineCount - 1;
			}

			return Enumerable.Range(start, end).Select(x => snapshot.GetLineFromLineNumber(x));
		}

		private string Replacer(string srcStr, int tabSize) {
			if (String.IsNullOrEmpty(srcStr)) return srcStr;
			string result = srcStr;

			// обработка начальных пробелов
			if (result[0] == ' ') {
				int spcStart = -1;
				int spcCount = 0;

				for (int i = 0; i < result.Length; i++) {
					if (result[i] == ' ' && spcStart == -1) spcStart = i;   // нашли первый пробел в строке
					if (result[i] != ' ') break;         // нашли последний пробел перед текстом
					if (result[i] == ' ') spcCount++;
				}

				if (spcStart != -1) {
					int tabsCount = CalcTabs(spcCount, tabSize);
					result = result.Remove(spcStart, spcCount);

					for (int i = 0; i < tabsCount; i++) {
						result = result.Insert(spcStart, "\t");
					}
				}
			}

			// пробелы перед комментарием
			Console.WriteLine("Check: " + result);
			int spcCount2 = 0;
			int commStart = result.IndexOf("//");
			Console.WriteLine("commStart: " + commStart);

			if (commStart > 0) {	// если перед комментарием пробел
				for (int i = commStart-1; i >= 0; i--) {
					if (result[i] != ' ') break;         // нашли первый пробел перед комментарием
					if (result[i] == ' ') spcCount2++;
				}
			}
			Console.WriteLine(spcCount2);

			if (spcCount2 > 0) {
				int tabsCount = CalcTabs(spcCount2, tabSize);
				int startInd = commStart - spcCount2;
				result = result.Remove(startInd, spcCount2);

				for (int i = 0; i < tabsCount; i++) {
					result = result.Insert(startInd, "\t");
				}

			}
			Console.WriteLine("Fixed: " + result);


			return result;
		}


		private int CalcTabs(int spcCount, int tabSize) {
			int num = spcCount / tabSize;
			int lastSpc = spcCount % tabSize;
			//if (lastSpc >= (tabSize/2)) {
			if (lastSpc > 0) num++;
			return num;
		}


	}
}
