# ConvertLearnToDoc

[![Build and deploy .NET Core app to Windows WebApp ConvertLearnToDocWeb](https://github.com/markjulmar/ConvertLearnToDoc/actions/workflows/ConvertLearnToDocWeb.yml/badge.svg)](https://github.com/markjulmar/ConvertLearnToDoc/actions/workflows/ConvertLearnToDocWeb.yml)

Tools to convert Learn module to Word doc and back

## Project structure

The GitHub repo has several related projects:

| Project | Description |
|---------|-------------|
| **ConvertLearnToDoc** | A simple CLI to convert a Learn module to a Word doc, or vice-versa. |
| **ConvertLearnToDocWeb** | A web portal version of the conversion tool. |
| **GrabAllLearnModules** | A CLI tool to walk a local clone of a Learn repo and create Word docs from each located module. |

In addition, there are four libraries used by the above projects.

| Library project | Description |
|-----------------|-------------|
| **Docx.Renderer.Markdown | A library to convert a .docx file to Markdown |
| **GenMarkdown.DocFX.Extensions | A library of [GenMarkdown](https://github.com/markjulmar/GenMarkdown) extensions to render DocFX extensions. |
| **LearnDocUtils** | The main conversion library. |
| **Markdig.Renderer.Docx** | A Markdig library to read a Markdig document and turn it into a .docx file. |

## Project dependencies

The project also depends on several NuGet packages:

| Package | Description |
|---------|-------------|
| **DxPlus** | A library to read/write .docx files. |
| **GenMarkdown** | A library to generate Markdown content. |
| **Markdig** | A markdown parsing library |
| **Microsoft.DocAsCode.MarkdigEngine.Extensions** | Extensions for Markdig and DocFX. |

## Converting a Learn module to a Word document

To try out the tools locally, clone the repository and navigate to the `ConvertLearnToDoc\src` project folder. Running the tool with no parameters will list the options:

```output
-i, --input        Required. Input file or folder.
-o, --output       Required. Output file or folder.
-z, --zipOutput    Zip output folder, defaults to false.
-r, --Repo         GitHub repo
-b, --Branch       GitHub branch, defaults to 'live'.
-t, --Token        GitHub access token.
-d, --Debug        Debug output, save temp files.
-p, --Pivot        Zone pivot to render to doc
-n, --Notebook     Convert notebooks into document
```

| Option | Description |
|--------|-------------|
| `-i` | Specifies a local Learn module folder, URL to a Learn module, or a local .docx file. |
| `-o` | Specifies a local folder to output a Learn module to, or a .docx filename. |
| `-z` | If supplied and converting from Learn to .docx, this will zip the generated folder. |
| `-r` | Repository to pull content from. If provided, the input parameter should a folder in this repo. |
| `-b` | Optional branch if content is not public. If provided, the input parameter should a folder in this repo. |
| `-t` | GitHub token if using a URL or GitHub folder. |
| `-d` | Debug - keeps all intermediary files. |
| `-p` | Zone pivot to render when going from Learn to a .docx. If not supplied, all pivots are rendered. |
| `-n` | If supplied, any notebooks in the module will be rendered in place. |
