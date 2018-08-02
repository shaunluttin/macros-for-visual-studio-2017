/// <reference path="C:\Users\bigfo\AppData\Local\Microsoft\VisualStudio\15.0_26cdf119\Macros\dte.js" />

//
// Work in Progress
//

Macro.InsertText('foo');

var sln = dte.Solution;
var slnName = dte.Solution.FullName;
Macro.InsertText('sln' + slnName);

//var objs = dte.Windows;
//var objs = dte.Documents;
var objs = dte.Solution.Projects;
for (var i = 1; i <= objs.Count; i++) {
    var o = objs.Item(i);

    dte.ActiveDocument.Selection.NewLine();
    Macro.InsertText('obj' + i);

    dte.ActiveDocument.Selection.NewLine();
    Macro.InsertText(o.Name);
}

Macro.InsertText('baz');
