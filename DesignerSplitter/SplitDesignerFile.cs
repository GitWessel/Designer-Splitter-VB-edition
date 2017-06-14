//------------------------------------------------------------------------------
// <copyright file="SplitDesignerFile.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Collections;
using System.ComponentModel.Design;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio.Shell;

namespace DesignerSplitter
{
	/// <summary>
	/// Command handler
	/// </summary>
	internal sealed class SplitDesignerFile
	{
		/// <summary>
		/// Command ID.
		/// </summary>
		public const int CommandId = 0x0100;

		/// <summary>
		/// Command menu group (command set GUID).
		/// </summary>
		public static readonly Guid CommandSet = new Guid("fab92d28-2657-4141-99e5-117b479c0b68");

		/// <summary>
		/// VS Package that provides this command, not null.
		/// </summary>
		private readonly Package package;

		/// <summary>
		/// Initializes a new instance of the <see cref="SplitDesignerFile"/> class.
		/// Adds our command handlers for menu (commands must exist in the command table file)
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		private SplitDesignerFile(Package package)
		{
            this.package = package ?? throw new ArgumentNullException("package");

			OleMenuCommandService commandService =
				ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
			if (commandService != null)
			{
				var menuCommandID = new CommandID(CommandSet, CommandId);
				var menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
				commandService.AddCommand(menuItem);
			}
		}

		/// <summary>
		/// Gets the instance of the command.
		/// </summary>
		public static SplitDesignerFile Instance { get; private set; }

		/// <summary>
		/// Gets the service provider from the owner package.
		/// </summary>
		private IServiceProvider ServiceProvider
		{
			get { return package; }
		}

		/// <summary>
		/// Initializes the singleton instance of the command.
		/// </summary>
		/// <param name="package">Owner package, not null.</param>
		public static void Initialize(Package package)
		{
			Instance = new SplitDesignerFile(package);
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
			ExtractDesignerFile();
		}

		private void ExtractDesignerFile()
		{
			var dte = (DTE) Package.GetGlobalService(typeof(DTE));

			var item = dte.SelectedItems.Item(1).ProjectItem;
			var filename = item.FileNames[1];
			var dir = Path.GetDirectoryName(filename);
			var bareName = Path.GetFileNameWithoutExtension(filename);
			var newItemPath = string.Format("{0}\\{1}.Designer.vb", dir, bareName);

			var usingStatementElements = FindUsingStatements(item.FileCodeModel.CodeElements);

			var usingStatements = new StringBuilder();
			foreach (var import in usingStatementElements)
			{
				usingStatements.AppendLine("Imports " + import.Namespace);
			}

			var codeClass = FindClass(item.FileCodeModel.CodeElements) as CodeClass;
			if (codeClass == null)
			{
				return;
			}

			// Mark class as partial
			((CodeClass2) codeClass).ClassKind = vsCMClassKind.vsCMClassKindPartialClass;
           
			var initComponentText = ExtractMember(codeClass.Members.Item("InitializeComponent"));
			var disposingText = ExtractMember(codeClass.Members.Item("Dispose"));
			var fieldDecls = ExtractWinFormsFields(codeClass);

			var toWrite = new StringBuilder();
			toWrite.AppendLine(usingStatements.ToString());
			toWrite.AppendLine();
			toWrite.AppendLine(string.Format("\tpublic partial class {0}", codeClass.Name));
			toWrite.AppendLine(EnsureEachNewLineStartsWith("\t\t", fieldDecls));
			toWrite.AppendLine();
			toWrite.AppendLine(EnsureEachNewLineStartsWith("\t\t", disposingText));
			toWrite.AppendLine();
			toWrite.AppendLine("\t\t#region \"Windows Form Designer generated code\"");
			toWrite.AppendLine();
			toWrite.AppendLine(EnsureEachNewLineStartsWith("\t\t", initComponentText));
			toWrite.AppendLine();
			toWrite.AppendLine("\t\t#End Region");
			toWrite.AppendLine();
            toWrite.AppendLine("End Class");

			File.WriteAllText(newItemPath, toWrite.ToString());

			var newProjItem = item.ProjectItems.AddFromFile(newItemPath);
			newProjItem.Open();
		}

		private string EnsureEachNewLineStartsWith(string prepend, string contents)
		{
			var toReturn = new StringBuilder();
			var lines = contents.Split(new[] {Environment.NewLine}, StringSplitOptions.None).ToArray();
			foreach (var line in lines)
			{
				if (line.StartsWith(prepend))
				{
					toReturn.AppendLine(line);
					continue;
				}
				toReturn.AppendLine(prepend + line);
			}
			return toReturn.ToString();
		}

		private CodeElement FindClass(IEnumerable codeElements)
		{
			foreach (CodeElement element in codeElements)
			{
				if (element.Kind == vsCMElement.vsCMElementClass)
				{
					return element;
				}
				if (element.Children.Count <= 0) continue;
				var cls = FindClass(element.Children);
				if (cls != null)
					return FindClass(element.Children);
			}
			return null;
		}

		private IEnumerable<CodeImport> FindUsingStatements(IEnumerable codeElements)
		{
			var toReturn = new List<CodeImport>();
			foreach (CodeElement element in codeElements)
			{
				Console.WriteLine(element.Kind);
				if (element is CodeImport)
				{
					toReturn.Add(element as CodeImport);
				}
				if (element.Children.Count <= 0) continue;
				var imports = FindUsingStatements(element.Children);
				if (imports != null && imports.Any())
				{
					toReturn.AddRange(imports);
				}
			}
			return toReturn;
		}

		private string ExtractMember(CodeElement element)
		{
			var memberStart = element.GetStartPoint().CreateEditPoint();
			var memberText = string.Empty;
			memberText += memberStart.GetText(element.GetEndPoint());
			memberStart.Delete(element.GetEndPoint());
			return memberText;
		}

		public string ExtractMember(CodeVariable variable)
		{
			var memberStart = variable.GetStartPoint().CreateEditPoint();
			var memberText = string.Empty;
			memberText += memberStart.GetText(variable.GetEndPoint());
			memberStart.Delete(variable.GetEndPoint());
			return memberText;
		}

		private string ExtractWinFormsFields(CodeClass codeClass)
		{
			var fieldsCode = new StringBuilder();
			foreach (var member in codeClass.Members)
			{
				var element = member as CodeElement;
				if (element == null) continue;

				if (element.Kind == vsCMElement.vsCMElementVariable)
				{
					var field = member as CodeVariable;
					if (field == null) continue;
					try
					{
						var fieldType = field.Type.CodeType;
						if (DetermineIfControl(fieldType))
						{
							fieldsCode.AppendLine(ExtractMember(field));
						}
					}
					catch (Exception e)
					{
						Console.WriteLine(e);
					}
				}
			}
			return fieldsCode.ToString();
		}

		private bool DetermineIfControl(CodeType fieldType)
		{
			if (fieldType.Namespace.FullName.StartsWith("System.Windows.Forms"))
				return true;

			if (fieldType.Namespace.FullName.StartsWith("DevExpress."))
				return true;

			if (fieldType.IsDerivedFrom["System.Windows.Forms.Control"])
				return true;

			if (fieldType.IsDerivedFrom["System.ComponentModel.IContainer"])
				return true;

			if (fieldType.IsDerivedFrom["System.ComponentModel.Container"])
				return true;

			return false;
		}
	}
}