---
title: Presentation.ExportAsFixedFormat3 method (PowerPoint)
keywords: vbapp10.chm583139
f1_keywords:
- vbapp10.chm583139
ms.assetid: 55a9c44e-e82d-4cb8-9b21-bd491087c1e9
ms.date: 07/10/2024
ms.localizationpriority: medium
---

# Presentation.ExportAsFixedFormat3 method (PowerPoint)

Publishes a copy of a Microsoft PowerPoint presentation as a file in a fixed format, either PDF or XPS.

## Syntax

_expression_.**ExportAsFixedFormat3** (_Path_, _FixedFormatType_, _Intent_, _FrameSlides_, _HandoutOrder_, _OutputType_, _PrintHiddenSlides_, _PrintRange_, _RangeType_, _SlideShowName_, _IncludeDocProperties_, _KeepIRMSettings_, _DocStructureTags_, _BitmapMissingFonts_, _UseISO19005_1_, _IncludeMarkup_, _ExternalExporter_, _Bookmarks_, _DocumentMarkup_, _PromotedHyperlinkShape_)

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
| _PrintRange_|Required|**PrintRange**|The slide range. Can be set to `Nothing` for all|
| _RangeType_|Optional|**PpPrintRangeType**|The type of slide range.|
| _SlideShowName_|Optional|**String**|The name of the slide show.|
| _IncludeDocProperties_|Optional|**Boolean**|Whether the document properties should also be exported. The default is **False**.|
| _KeepIRMSettings_|Optional|**Boolean**|Whether the IRM settings should also be exported.</br></br>If _FixedFormatType_ is _PpFixedFormatTypePDF_, this flag determines if labels and IRM settings should be exported.</br></br>The default is **True**.|
| _DocStructureTags_|Optional|**Boolean**|Whether to include document structure tags to improve document accessibility. The default is **True**.|
| _BitmapMissingFonts_|Optional|**Boolean**|Whether to include a bitmap of the text. The default is **True**.|
| _UseISO19005_1_|Optional|**Boolean**|Whether the resulting document is compliant with ISO 19005-1 (PDF/A). The default is **False**.|
| _IncludeMarkup_|Optional|**Boolean**|Whether the resulting document should include associated pen marks.|
| _Bookmarks_|Optional|**Boolean**|Whether bookmarks for each section and slide should be included in the exported document. When using this option, external exporters should not add their own bookmarks for sections or slides. The default is **True**.|
| _DocumentMarkup_|Optional|**Boolean**|Whether the **Document** tag should be included in the document structure tags. When using this option, external exporters should not add their own **Document** tag. The default is **True**.|
| _PromotedHyperlinkShape_|Optional|**Boolean**|Whether hyperlinks should be promoted to siblings of objects rather than nested within objects in document structure tags. Transparent text elements with alpha of 0 are included for the hyperlinks and external exporters should respect the alpha value so that they are not visible in the document. The default is **True**.|
| _ExternalExporter_|Optional|**Variant**|A pointer to an Office add-in that implements the **IMsoDocExporter** COM interface and allows calls to an alternate implementation of code. The default is a null pointer.|

## Return value

**VOID**

## Remarks

The _KeepIRMSettings_ parameter behaves specially for PDF. It controls the retention of both labels and encryption to the output file. For more information, see [Manage sensitivity labels in Office apps](/microsoft-365/compliance/sensitivity-labels-office-apps?view=o365-worldwide#pdf-support&preserve-view=true).

If the presentation is not fully downloaded, this method fails and an error occurs. For more information about the Partial Documents, see [Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md).

PrintRange is a required parameter, but may be set to `Nothing`. If not supplied, the call will fail with a `Type mismatch`

## Requirements

Microsoft 365 Version 2408 (Build 17928.xxxxx)

### See also
[Manage sensitivity labels in Office apps](/microsoft-365/compliance/sensitivity-labels-office-apps?view=o365-worldwide#pdf-support&preserve-view=true)

[Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
