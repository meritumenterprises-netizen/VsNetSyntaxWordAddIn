# Microsoft Word Add-In for Syntax Coloring C#, JavaScript, TypeScript, HTML, Angular, JSON and XML

The add-in creates a VS.NET Syntax tab in the Word ribbon for formatting code snippets in C#, HTML, TypeScript, JavaScript, Angular, JSON, and XML. It uses Roslyn for C# and Serenity.TypeScript for TypeScript to parse and highlight syntax. The add-in is built with VSTO targeting .NET Framework 4.8 and does not require Office.js or a web server.

The add-in looks like below:

Language dropdown:

<table>
<tr>
<th width="550" align="center">
<b>Language</b>
</th>
</tr>
<tr>
<td>
C# 12
</td>
</tr>
<tr>
<td>
JavaScript
</td>
</tr>
<tr>
<td>
TypeScript
</td>
</tr>
<tr>
<td>
HTML
</td>
</tr>
<tr>
<td>
Angular
</td>
</tr>
<tr><td>JSON</td></tr>
<tr><td>XML</td></tr>
</table>

Themes:

<table>
  <tr>
    <th width="150">Theme</th>
    <th width="400">Description</th>
  </tr>
  <tr>
    <td>Light</td>
    <td>Light theme from VS.NET 2026</td>
  </tr>
  <tr>
    <td>Dark</td>
    <td>Dark theme from VS.NET 2026</td>
  </tr>
</table>



<table>
  <tr>
    <th width="150">Background</th>
    <th width="400">Description</th>
  </tr>
  <tr>
    <td>Normal</td>
    <td>No change in background color</td>
  </tr>
  <tr>
    <td>Red</td>
    <td>Light-red background (used for incorrect code)</td>
  </tr>
  <tr>
    <td>Green</td>
    <td>Light-green background (used for correct code</td>
  </tr>
</table>

Runtime formatting is performed by either selections from the VS.NET Syntax Ribbon tab, or by using keyboard shortcuts:

<table>
  <tr>
    <th width="100">Keyboard shortcut</th>
    <th width="400">Meaning</th>
  </tr>
  <tr>
    <td>Ctrl+Alt+Shift+C</td>
    <td>Format text as C# 12 syntax, use background color from the VS.NET Syntax Ribbon tab</td>
  </tr>
  <tr>
    <td>Ctrl+Alt+Shift+T</td>
    <td>Format text as TypeScript syntax, use background color from the VS.NET Syntax Ribbon tab</td>
  </tr>
  <tr>
    <td>Ctrl+Alt+Shift+J</td>
    <td>Format text as JavaScript syntax, use background color from the VS.NET Syntax Ribbon tab</td>
  </tr>
  <tr>
    <td>Ctrl+Alt+Shift+H</td>
    <td>Format text as HTML syntax, use background color from the VS.NET Syntax Ribbon tab</td>
  </tr>
  <tr>
    <td>Ctrl+Alt+Shift+A</td>
    <td>Format text as Angular syntax, use background color from the VS.NET Syntax Ribbon tab</td>
  </tr>
  <tr>
    <td>Ctrl+Shift+Alt+S</td>
    <td>Format text as JSON syntax, use background color from the VS.NET Syntax Ribbon tab</td>
  </tr>
  <tr>
    <td>Ctrl+Shift+Alt+X</td>
    <td>Format text as XML syntax, use background color from the VS.NET Syntax Ribbon tab</td>
  </tr>
</table>

# VS.NET Syntax Word Add-in

Native C# VSTO add-in for desktop Microsoft Word on Windows, targeting .NET Framework 4.8. No Office.js, browser, or web server is required.

## Build and Run

Install Visual Studio with the Office/SharePoint development workload and desktop Word.

1. Run `pwsh -NoProfile -File tools/Run.ps1` to generate a local development signing certificate, restore packages, build signed VSTO manifests, register the add-in for your Windows account, and launch a new Word instance.
2. For debugging, open `VsNetSyntaxWordAddIn.sln` from its current folder in Visual Studio. After moving the project, reopen this solution so Visual Studio uses the new paths.
3. Press F5 to register the development add-in and launch Word.
4. Select plain code text. In the **VS.NET Syntax** tab, choose **C#**, **HTML**, **TypeScript**, **JavaScript**, **Angular**, **JSON**, or **XML** from **Language**. Changing the language applies formatting immediately.
5. Use **Apply** to format another selection with the current language.

The **Theme** dropdown offers **Light** (the default) and **Dark** for language text colors. The **Background** dropdown offers **Normal** (the default, leaves existing background formatting untouched), **Red** (`#FED9B4`), and **Green** (`#EDF5E7`), with rectangular color swatches in the menu and on the closed control for the selected value. Office places dropdown images to the left; right-side images are not supported by this native control. Changing either dropdown immediately reformats the selection. Apply and keyboard shortcuts use both current settings. Background colors apply only to selected code, leaving the document page color unchanged.

Formatting modifies character properties in place, preserving text, tabs, line breaks, and the selection. Changes are grouped into one Undo action. C# and TypeScript class, method, and property names are bold; comments, strings, and keywords retain normal weight. C# resolves references available in the selected snippet with Roslyn; TypeScript uses the Serenity.TypeScript parser for declarations and syntactically recognizable member references. Neither has the full source project's symbol context. HTML and JavaScript are unchanged. Reapplying formatting resets obsolete bold styling. The formatter uses Consolas 10 pt and rejects protected documents and selections containing cell markers, fields, inline objects, or content controls.

To build without registering or launching Word, run `pwsh -NoProfile -File tools/Build.ps1`. To build and register only, add `-Register`. The build script locates Visual Studio by its installed MSBuild and OfficeTools targets, avoiding installer component identifiers that differ between installations.

Both left and right modifier keys work. Holding a shortcut applies it once per key press. The ribbon language changes to match the shortcut. The loaded Word add-in owns these native Windows hotkeys: it registers them while a Word document window is active and releases them when focus leaves Word, the add-in unloads, or Word closes. The old Windows-key shortcuts are no longer registered.

C# uses Microsoft's Roslyn parser (`Microsoft.CodeAnalysis.CSharp` 4.8.0), configured for **C# 12**. It recognizes modern syntax including records, required/init members, primary constructors, collection expressions, raw and UTF-8 strings, and interpolated expressions. Contextual keywords are colored according to their parsed roles, not a word list. Incomplete snippets are handled by Roslyn's syntax recovery. This is syntax highlighting, not compilation or full project-aware semantic classification; symbols and external types are not resolved.

ColorCode.Core 2.0.15 supplies the HTML, TypeScript, JavaScript, JSON, and XML grammars and the DefaultLight/DefaultDark palettes for all languages. JSON colors distinguish keys, strings, numbers, and true/false/null constants. XML colors distinguish element names, attributes, values, delimiters, comments, entities, and CDATA. Exact Visual Studio 2026 theme and semantic-color parity is not claimed.

## Deployment

The build produces `bin/Debug/VsNetSyntaxWordAddIn.vsto`. For distribution, use Visual Studio Publish with your own trusted code-signing certificate. The generated development certificate is for local testing only. Do not commit `DevelopmentSigning.pfx` or `DevelopmentSigning.props`.
