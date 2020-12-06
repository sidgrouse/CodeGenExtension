using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace CodeGenExtension
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class CreateConstructorCommand
    {
        private readonly Regex _propRegex = new Regex(
            @"^[\s]+public[\s]+(?<type>[\w]+)[\s]+(?<name>[\w]+)[\s]+{[\s+]get",
            RegexOptions.Multiline | RegexOptions.Compiled);
        private readonly Regex _classNameRegex = new Regex(
            @"^[\s]+public[\s]class[\s]+(?<class>[\w]+)[\s]+{",
            RegexOptions.Multiline | RegexOptions.Compiled);
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("ccef335d-78a8-43be-949f-726fdb9452ea");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="CreateConstructorCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private CreateConstructorCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static CreateConstructorCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
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
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in CreateConstructorCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new CreateConstructorCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            var currentTextFile = GetActiveText();
            var properties = GetProperties(currentTextFile);
            var (fileName, offset) = GetClassNameAndOffset(currentTextFile);
            var constructorText = FormConstructor(properties, fileName);
            InsertConstructor(constructorText, offset);
        }

        private List<Property> GetProperties(string classText)
        {
            var propertyMatches = _propRegex.Matches(classText);
            var result = new List<Property>();

            for (int i = 0; i < propertyMatches.Count; i++)
            {
                result.Add(new Property
                {
                    Name = propertyMatches[i].Groups["name"].Value,
                    Type = propertyMatches[i].Groups["type"].Value
                });
            }

            return result;
        }

        private (string, int) GetClassNameAndOffset(string classText)
        {
            var match = _classNameRegex.Match(classText.Replace("\r\n", "\n"));
            var offset = match.Index + match.Value.Length + 1;
            return (match.Groups["class"].Value, offset);
        }

        private static string GetActiveText()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            DTE dte = Package.GetGlobalService(typeof(SDTE)) as DTE;
            var doc = dte.ActiveDocument.Object() as TextDocument;
            return doc?.CreateEditPoint(doc.StartPoint)?.GetText(doc.EndPoint) ?? string.Empty;
        }

        private static string FormConstructor(List<Property> properties, string className)
        {
            var secondIndentation = "\t\t";
            var thirdIndentation = "\t\t\t";
            var lastPropertyNumber = properties.Count - 1;

            if (properties.Count == 0)
            {
                return $"\r\n{secondIndentation}public {className}()\r\n{thirdIndentation}{{\r\n{thirdIndentation}}}";
            }

            var sb = new StringBuilder($"\r\n{secondIndentation}public {className}(\r\n");
            foreach (var property in properties.Take(properties.Count - 1))
            {
                sb.AppendLine($"{thirdIndentation}{property.Type} {property.GetParameterName()},");
            }
            sb.AppendLine($"{thirdIndentation}{properties[lastPropertyNumber].Type} {properties[lastPropertyNumber].GetParameterName()})\r\n{secondIndentation}{{");

            foreach (var property in properties)
            {
                sb.AppendLine($"{thirdIndentation}{property.Name} = {property.GetParameterName()};");
            }
            sb.AppendLine($"{secondIndentation}}}");

            return sb.ToString();
        }

        private static void InsertConstructor(string constructor, int offset)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            DTE dte = Package.GetGlobalService(typeof(SDTE)) as DTE;
            var doc = dte.ActiveDocument.Object() as TextDocument;
            var point = doc?.CreateEditPoint();
            point.MoveToAbsoluteOffset(offset);
            point.Insert(constructor);
        }

        private class Property
        {
            public string Type { get; set; }

            public string Name { get; set; }

            public string GetParameterName()
            {
                return Name[0].ToString().ToLower() + Name.Remove(0, 1);
            }
        }
    }
}
