# ConvertLearnToDoc

[![Build and deploy .NET Core app to Windows WebApp ConvertLearnToDocWeb](https://github.com/markjulmar/ConvertLearnToDoc/actions/workflows/ConvertLearnToDocWeb.yml/badge.svg)](https://github.com/markjulmar/ConvertLearnToDoc/actions/workflows/ConvertLearnToDocWeb.yml)

Tools to convert Learn module to Word doc and back

## Project structure

The GitHub repo has several related projects:

| Project | Description |
|---------|-------------|
| [Blazor](https://github.com/markjulmar/ConvertLearnToDoc/tree/main/Blazor/) | A Blazor client/server version of the conversion tool.
| [ConvertLearnToDocWeb](https://github.com/markjulmar/ConvertLearnToDoc/tree/main/src/ConvertLearnToDocWeb) | A web portal version of the conversion tool. This is no longer published and has been replaced with the Blazor version. |
| [ConvertToDocAzureFunctions](https://github.com/markjulmar/ConvertLearnToDoc/tree/main/src/ConvertLearnToDoc.AzureFunctions) | Azure functions to perform the document conversions, used by the above web project. This is no longer published and has been replaced with the Blazor version. |
| [ConvertAll](https://github.com/markjulmar/ConvertLearnToDoc/tree/main/src/ConvertAll) | A CLI tool to walk a local clone of a Learn repo and create Word docs from each located module. |
| [ConvertDocx](https://github.com/markjulmar/ConvertLearnToDoc/tree/main/src/ConvertDocx) | A CLI tool to convert a single Learn module or Docs page to a Word doc, or vice-versa. It can take a URL, GitHub details, or a local folder/file. |

In addition, there are four libraries used by the above projects.

| Library project | Description |
|-----------------|-------------|
| [Docx.Renderer.Markdown](https://github.com/markjulmar/ConvertLearnToDoc/tree/main/lib/Docx.Renderer.Markdown) | A library to convert a .docx file to Markdown |
| [GenMarkdown.DocFX.Extensions](https://github.com/markjulmar/ConvertLearnToDoc/tree/main/lib/GenMarkdown.DocFx.Extensions) | A library of [GenMarkdown](https://github.com/markjulmar/GenMarkdown) extensions to render DocFX extensions. |
| [LearnDocUtils](https://github.com/markjulmar/ConvertLearnToDoc/tree/main/lib/LearnDocUtils) | The main conversion library. |
| [Markdig.Renderer.Docx](https://github.com/markjulmar/ConvertLearnToDoc/tree/main/lib/Markdig.Renderer.Docx) | A Markdig library to read a Markdig document and turn it into a .docx file. |

## Project dependencies

The project also depends on several NuGet packages:

| Package | Description |
|---------|-------------|
| [DxPlus](https://github.com/markjulmar/dxplus) | A library to read/write .docx files. |
| [GenMarkdown](https://github.com/markjulmar/genmarkdown) | A library to generate Markdown content. |
| [Markdig](https://github.com/xoofx/markdig) | A markdown parsing library |
| [MSLearnRepos](https://www.nuget.org/packages/julmar.mslearnrepos) | A .NET library to work with GitHub and the Learn repo structure |
| [Microsoft.DocAsCode.MarkdigEngine.Extensions](https://www.nuget.org/packages/Microsoft.DocAsCode.MarkdigEngine.Extensions) | Extensions for Markdig and DocFX. |

## Converting a Learn module to a Word document

To try out the tools locally, clone the repository and navigate to the `src\ConvertDocx` project folder. Running the tool with no parameters will list the options:

```output
-i, --input           Required. Input file or folder.
-o, --output          Required. Output file or folder.
-s, --singlePage      Output should be a single page (Markdown file).
-g, --Organization    GitHub organization
-r, --Repo            GitHub repo
-b, --Branch          GitHub branch, defaults to 'live'
-t, --Token           GitHub access token
-d, --Debug           Debug output, save temp files
-p, --Pivot           Zone pivot to render to doc, defaults to all
-z, --zipOutput       Zip output folder, defaults to false
-n, --Notebook        Convert notebooks into document, only used on MS Learn content
--help                Display help.
--version             Display version information.
```

| Option | Description |
|--------|-------------|
| `-i` | Specifies a local Learn module folder or docs Markdown page, URL to a Learn module/docs conceptual page, or a local .docx file. |
| `-o` | Specifies a local folder or file to output a Learn module/docs page to, or a .docx filename. |
| `-z` | If supplied and converting from Learn to .docx, this will zip the generated folder. |
| `-g` | GitHub organization to get content from. This allows a fork of MicrosoftDocs to be used. |
| `-r` | Repository to pull content from. If provided, the input parameter should a folder in this repo. |
| `-b` | Optional branch if content is not public. If provided, the input parameter should a folder in this repo. |
| `-t` | GitHub token if using a URL or GitHub folder. |
| `-d` | Debug - keeps all intermediary files. |
| `-p` | Zone pivot to render when going from Learn to a .docx. If not supplied, all pivots are rendered. |
| `-n` | If supplied, any notebooks in the module will be rendered in place. |
| `-s` | Indicates to render to a single page. This is only necessary if the input is a Word doc and the output filename does not indicate it should be a Markdown file. |
