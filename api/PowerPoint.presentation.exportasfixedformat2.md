---
title: Presentation.ExportAsFixedFormat2 method (PowerPoint)
keywords: vbapp10.chm583126
f1_keywords:
- vbapp10.chm583126
ms.assetid: b1101e58-e6a8-9dd4-7071-1325ba71edb1
ms.date: 06/08/2017
ms.prod: powerpoint
ms.localizationpriority: medium
---

# Presentation.ExportAsFixedFormat2 method (PowerPoint)

Publishes a copy of a Microsoft PowerPoint presentation as a file in a fixed format, either PDF or XPS.

## Syntax

_expression_.**ExportAsFixedFormat2** (_Path_, _FixedFormatType_, _Intent_, _FrameSlides_, _HandoutOrder_, _OutputType_, _PrintHiddenSlides_, _PrintRange_, _RangeType_, _SlideShowName_, _IncludeDocProperties_, _KeepIRMSettings_, _DocStructureTags_, _BitmapMissingFonts_, _UseISO19005_1_, _IncludeMarkup_, _ExternalExporter_)

_expression_ A variable that represents a **[Presentation](PowerPoint.Presentation.md)** object.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Path_|Required|**String**|The path for the export.|
| _FixedFormatType_|Required|**PpFixedFormatType**|The format to which the slides should be exported.|
| _Intent_|Optional|**PpFixedFormatIntent**|The purpose of the export.|
| _FrameSlides_|Optional|**MsoTriState**|Whether the slides to be exported should be bordered by a frame.|
| _HandoutOrder_|Optional|**PpPrintHandoutOrder**|The order in which the handout should be printed.|
| _OutputType_|Optional|**PpPrintOutputType**|The type of output.|
| _PrintHiddenSlides_|Optional|**MsoTriState**|Whether to print hidden slides.|
| _PrintRange_|Optional|**PrintRange**|The slide range.|
| _RangeType_|Optional|**PpPrintRangeType**|The type of slide range.|
| _SlideShowName_|Optional|**String**|The name of the slide show.|
| _IncludeDocProperties_|Optional|**Boolean**|Whether the document properties should also be exported. The default is **False**.|
| _KeepIRMSettings_|Optional|**Boolean**|Whether the IRM settings should also be exported.</br></br>If _FixedFormatType_ is _PpFixedFormatTypePDF_, this flag determines if labels and IRM settings should be exported.</br></br>The default is **True**.|
| _DocStructureTags_|Optional|**Boolean**|Whether to include document structure tags to improve document accessibility. The default is **True**.|
| _BitmapMissingFonts_|Optional|**Boolean**|Whether to include a bitmap of the text. The default is **True**.|
| _UseISO19005_1_|Optional|**Boolean**|Whether the resulting document is compliant with ISO 19005-1 (PDF/A). The default is **False**.|
| _IncludeMarkup_|Optional|**Boolean**|Whether the resulting document should include associated pen marks.|
| _ExternalExporter_|Optional|**Variant**|A pointer to an Office add-in that implements the **IMsoDocExporter** COM interface and allows calls to an alternate implementation of code. The default is a null pointer.|
| _Path_|Required|**String**||
| _FixedFormatType_|Required|PPFIXEDFORMATTYPE||
| _Intent_|Optional|PPFIXEDFORMATINTENT||
| _FrameSlides_|Optional|unknown||
| _HandoutOrder_|Optional|PPPRINTHANDOUTORDER||
| _OutputType_|Optional|PPPRINTOUTPUTTYPE||
| _PrintHiddenSlides_|Optional|unknown||
| _PrintRange_|Optional|PRINTRANGE||
| _RangeType_|Optional|PPPRINTRANGETYPE||
| _SlideShowName_|Optional|**String**||
| _IncludeDocProperties_|Optional|BOOL||
| _KeepIRMSettings_|Optional|BOOL||
| _DocStructureTags_|Optional|BOOL||
| _BitmapMissingFonts_|Optional|BOOL||
| _UseISO19005_1_|Optional|BOOL||
| _IncludeMarkup_|Optional|BOOL||
| _ExternalExporter_|Optional|**Variant**||

## Return value

**VOID**

## Remarks

The _KeepIRMSettings_ parameter behaves specially for PDF. It controls the retention of both labels and encryption to the output file. For more information, see [Manage sensitivity labels in Office apps](/microsoft-365/compliance/sensitivity-labels-office-apps?view=o365-worldwide#pdf-support&preserve-view=true).

Due to the interaction of partner add-ins creating PDFs in Office with encryption, Office will default the _KeepIRMSettings_ flag to **FALSE** until _second RMID_ releases. 

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
