using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using EnvDTE;
using EnvDTE80;
using System.Collections.Generic;

namespace FindGroupBy
{
	/// <summary>
	/// Command handler
	/// </summary>
	internal sealed class FindGroupByFunction
	{
		/// <summary>
		/// Command ID.
		/// </summary>
		public const int CommandId = 0x0100;

		/// <summary>
		/// Command menu group (command set GUID).
		/// </summary>
		public static readonly Guid CommandSet = new Guid("3feb6580-a577-4d14-aa41-987074aa55c9");

		private const string nofunction = "no function";

		/// <summary>
		/// VS Package that provides this command, not null.
		/// </summary>
		private readonly Package package;

		private Dictionary<string, List<string>> dict;

		/// <summary>
		/// Initializes a new instance of the <see cref="FindGroupByFunction"/> class.
		/// Adds our command handlers for menu (commands must exist in the command table file)
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		private FindGroupByFunction(Package package)
		{
			if (package == null)
			{
				throw new ArgumentNullException("package");
			}

			this.package = package;

			OleMenuCommandService commandService = this.ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
			if (commandService != null)
			{
				var menuCommandID = new CommandID(CommandSet, CommandId);
				var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
				commandService.AddCommand(menuItem);
			}

			DTE2 dte = ServiceProvider.GetService(typeof(EnvDTE.DTE)) as DTE2;
			FindEvents e = dte.Events.FindEvents;
			e.FindDone += Fe_FindDone;
		}

		private void Fe_FindDone(vsFindResult findResult, bool Cancelled)
		{
			dict = new Dictionary<string, List<string>>();

			DTE2 dte = ServiceProvider.GetService(typeof(EnvDTE.DTE)) as DTE2;

			Window window;
			TextSelection textSelection;
			TextPoint textSelectionPointSaved;
			OutputWindowPane outputWindowPane;
			EnvDTE.TextDocument textDocument;
			int lastFoundAt = 0;

			textDocument = (TextDocument)dte.ActiveDocument.Object("");
			textSelection = textDocument.Selection;
			window = dte.ActiveDocument.Windows.Item(1);
			
			// Set up output window pane and loop until no more matches.
			outputWindowPane = GetOutputWindowPane(dte, "Matching Lines");
			textSelection.StartOfDocument();
			textSelectionPointSaved = textSelection.ActivePoint.CreateEditPoint();

			// GetOutputWindowPane activates Output Window, so re-activate our window.
			window.Activate();
			outputWindowPane.Clear();

			while (findResult == vsFindResult.vsFindResultFound)
			{
				if (textSelection.AnchorPoint.Line <= lastFoundAt)
					break;

				textSelection.SelectLine();

				string functionName = nofunction;
				EnvDTE.CodeFunction func = textSelection.AnchorPoint.CodeElement[vsCMElement.vsCMElementFunction] as EnvDTE.CodeFunction;
				if (func != null) functionName = func.FullName;

				if (!dict.ContainsKey(functionName))
				{
					dict.Add(functionName, new List<string>());
				}

				dict[functionName].Add(textSelection.Text);

				lastFoundAt = textSelection.AnchorPoint.Line;
				textSelection.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstColumn);
				findResult = dte.Find.Execute();
			}

			// Restore caret to location before invoking this command.
			textSelection.MoveToPoint(textSelectionPointSaved);

			PrintAll(outputWindowPane);
		}

		void m_findEvents_FindDone(EnvDTE.vsFindResult Result, bool Cancelled)
		{
			var dte = (EnvDTE.DTE)ServiceProvider.GetService(typeof(EnvDTE.DTE));
			// Get search term, window location, etc...;
			var x = dte.Find.FindWhat;
			var guid = dte.Find.ResultsLocation == vsFindResultsLocation.vsFindResults1 ?
					"{0F887920-C2B6-11D2-9375-0080C747D9A0}" : "{0F887921-C2B6-11D2-9375-0080C747D9A0}";

			var findWindow = dte.Windows.Item(guid);
			var selection = findWindow.Selection as TextSelection;
			// Get search text results;
			var endPoint = selection.AnchorPoint.CreateEditPoint();
			endPoint.EndOfDocument();
			var text = endPoint.GetLines(1, endPoint.Line);
		}

		/// <summary>
		/// Gets the instance of the command.
		/// </summary>
		public static FindGroupByFunction Instance
		{
			get;
			private set;
		}

		/// <summary>
		/// Gets the service provider from the owner package.
		/// </summary>
		private IServiceProvider ServiceProvider
		{
			get
			{
				return this.package;
			}
		}

		/// <summary>
		/// Initializes the singleton instance of the command.
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		public static void Initialize(Package package)
		{
			Instance = new FindGroupByFunction(package);
		}

		/// <summary>
		/// This function is the callback used to execute the command when the menu item is clicked.
		/// See the constructor to see how the menu item is associated with this function using
		/// OleMenuCommandService service and MenuCommand class.
		/// </summary>
		/// <param name="sender">Event sender.</param>
		/// <param name="e">Event args.</param>
		private void MenuItemCallback(object sender, EventArgs e)
		{
			dict = new Dictionary<string, List<string>>();

			DTE2 dte = ServiceProvider.GetService(typeof(EnvDTE.DTE)) as DTE2;

			Window window;
			TextSelection textSelection;
			TextPoint textSelectionPointSaved;
			OutputWindowPane outputWindowPane;
			EnvDTE.vsFindResult findResult;
			EnvDTE.TextDocument textDocument;
			int lastFoundAt = 0;

			textDocument = (TextDocument)dte.ActiveDocument.Object("");
			textSelection = textDocument.Selection;
			window = dte.ActiveDocument.Windows.Item(1);
			PrepareDefaultFind(dte, "List Matching Lines");

			// Set up output window pane and loop until no more matches.
			outputWindowPane = GetOutputWindowPane(dte, "Matching Lines");
			textSelection.StartOfDocument();
			textSelectionPointSaved = textSelection.ActivePoint.CreateEditPoint();

			// GetOutputWindowPane activates Output Window, so re-activate our window.
			window.Activate();
			outputWindowPane.Clear();
			((EnvDTE80.Find2)dte.Find).WaitForFindToComplete = true;

			findResult = dte.Find.Execute();
			while (findResult == vsFindResult.vsFindResultFound)
			{
				if (textSelection.AnchorPoint.Line <= lastFoundAt)
					break;

				textSelection.SelectLine();

				string functionName = nofunction;
				EnvDTE.CodeFunction func = textSelection.AnchorPoint.CodeElement[vsCMElement.vsCMElementFunction] as EnvDTE.CodeFunction;
				if (func != null) functionName = func.FullName;

				if (!dict.ContainsKey(functionName))
				{
					dict.Add(functionName, new List<string>());
				}

				dict[functionName].Add(textSelection.Text);

				lastFoundAt = textSelection.AnchorPoint.Line;
				textSelection.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstColumn);
				findResult = dte.Find.Execute();
			}

			// Restore caret to location before invoking this command.
			textSelection.MoveToPoint(textSelectionPointSaved);

			PrintAll(outputWindowPane);
		}

		public void PrintAll(OutputWindowPane outputWindowPane)
		{
			if (dict.Count == 0)
			{
				outputWindowPane.OutputString("No Results\r\n");
				return;
			}

			// print results not inside a function first
			if (dict.ContainsKey(nofunction))
			{
				var lista = dict[nofunction];
				if (lista.Count>0)
				{
					outputWindowPane.OutputString("#region " + nofunction + ":\r\n");

					foreach (var line in lista)
						outputWindowPane.OutputString(line);

					outputWindowPane.OutputString("\r\n#endregion\r\n");
				}

				dict.Remove(nofunction);
			}

			foreach (var func in dict.Keys)
			{
				var lista = dict[func];
				if (lista.Count > 0)
				{
					outputWindowPane.OutputString("#region " + func + ":\r\n");

					foreach (var line in lista)
						outputWindowPane.OutputString(line);

					outputWindowPane.OutputString("\r\n#endregion\r\n");
				}
			}
		}

		public string PrepareDefaultFind(DTE2 DTE, string prompt)
		{
			string what;
			DTE.Find.Action = vsFindAction.vsFindActionFind;
			DTE.Find.MatchCase = false;
			DTE.Find.MatchWholeWord = true;
			//DTE.Find.Target = vsFindTarget.vsFindTargetCurrentDocument;
			//DTE.Find.ResultsLocation = vsFindResultsLocation.vsFindResults1;
			DTE.Find.Backwards = false;
			DTE.Find.MatchInHiddenText = true;
			DTE.Find.PatternSyntax = vsFindPatternSyntax.vsFindPatternSyntaxLiteral;
			what = Microsoft.VisualBasic.Interaction.InputBox(prompt);

			if (what != "") DTE.Find.FindWhat = what;

			return what;
		}

		public OutputWindowPane GetOutputWindowPane(DTE2 dte, string name)
		{
			// Create a tool window reference for the Output window
			// and window pane.
			OutputWindow ow = dte.ToolWindows.OutputWindow;
			OutputWindowPane owp;

			try
			{
				owp = ow.OutputWindowPanes.Item(name);
			}
			catch
			{
				owp = ow.OutputWindowPanes.Add(name);
			}

			owp.Activate();
			return owp;
		}
	}
}
