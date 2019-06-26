---
title: Project.ExportAsFixedFormat method (Project)
keywords: vbapj.chm132843
f1_keywords:
- vbapj.chm132843
ms.prod: project-server
api_name:
- Project.Project.ExportAsFixedFormat
ms.assetid: ee217506-bcc5-a514-0c32-ba6402ac07f2
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.ExportAsFixedFormat method (Project)

Exports the active project as a document in a custom PDF or XPS format.


## Syntax

_expression_.**ExportAsFixedFormat** (_FileName_, _FileType_, _IncludeDocumentProperties_, _IncludeDocumentMarkup_, _ArchiveFormat_, _FromDate_, _ToDate_, _FixedFormatExtClassPtr_)

 _expression_ An expression that returns a **[Project](project.project.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|Specifies the file name of the exported document. The default value is the name of the active project as a PDF file.|
| _FileType_|Optional|**PjDocExportType**|Specifies whether to export the project as a PDF or an XPS document. The default value is  **pjPDF** (0).|
| _IncludeDocumentProperties_|Optional|**Boolean**|If  **True**, the last page of the exported document includes some document properties. The default value is **True**.|
| _IncludeDocumentMarkup_|Optional|**Boolean**|If  **True**, the last page of the exported document includes a legend of the symbols shown in the view. The default value is **True**.|
| _ArchiveFormat_|Optional|**Boolean**|If  **True**, exports a PDF document in the ISO 19500-1 compliant (PDF/A) format. The default value is **False**.|
| _FromDate_|Optional|**Variant**|The start date of the range of dates to publish. The default value is the project start date.|
| _ToDate_|Optional|**Variant**|The end date of the range of dates to publish. The default value is the project end date.|
| _FixedFormatExtClassPtr_|Optional|**Variant**|Pointer to a custom class in an add-in that implements the  **IMsoDocExporter** COM interface that allows calls to an alternate implementation of code for the document format. The default is a null pointer.|

## Return value

 **Nothing**


## Remarks

The  **ExportAsFixedFormat** method is similar to the **[DocumentExport](Project.Application.DocumentExport.md)** method, except the _FileName_ parameter is required and the optional _FixedFormatExtClassPtr_ parameter is a pointer to a user-defined class that creates a custom PDF or XPS format.


## Example

If the active project shows a Network Diagram view, the following example creates an XPS document named TestProject.xps. When you open the file in the  **XPS Viewer** application, the last page includes document properties and a legend that shows the PERT chart symbols.


```vb
ExportAsFixedFormat FileName:="TestProject.xps", FileType:=pjXPS
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]