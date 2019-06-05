---
title: Document.ExportAsFixedFormat method (Publisher)
keywords: vbapb10.chm196758
f1_keywords:
- vbapb10.chm196758
ms.prod: publisher
api_name:
- Publisher.Document.ExportAsFixedFormat
ms.assetid: 8bb5b64f-57b2-cf87-344c-be1e2741a59c
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.ExportAsFixedFormat method (Publisher)

Saves a Microsoft Publisher publication in PDF or XPS format. The conversion readies the document to be sent to commercial presses, to copy shops, for desktop printing, or for electronic distribution.


## Syntax

_expression_.**ExportAsFixedFormat** (_Format_, _FileName_, _Intent_, _IncludeDocumentProperties_, _ColorDownsampleTarget_, _ColorDownsampleThreshold_, _OneBitDownsampleTarget_, _OneBitDownsampleThreshold_, _From_, _To_, _Copies_, _Collate_, _PrintStyle_, _DocStructureTags_, _BitmapMissingFonts_, _UseISO19005\_1_, _ExternalExporter_)

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Format_|Required| **[PbFixedFormatType](publisher.pbfixedformattype.md)** |The format in which you want to export the publication. Can be one of the **PbFixedFormatType** constants.|
|_FileName_|Required| **String**|The file name for the exported file.|
|_Intent_|Optional| **[PbFixedFormatIntent](publisher.pbfixedformatintent.md)** |The output quality of the exported file. Can be one of the **PbFixedFormatIntent** constants.|
|_IncludeDocumentProperties_|Optional| **Boolean**| **True** if you want to save the document properties with the PDF file.|
|_ColorDownsampleTarget_|Optional| **Long**|The target for down-sampling of colored images. Measured in dots per inch. Must be greater than 96. |
|_ColorDownsampleThreshold_|Optional| **Long**|The threshold at or above which an image is down-sampled to the _ColorDownsampleTarget_ level.|
|_OneBitDownsampleTarget_|Optional| **Long**|The target for down-sampling of one-bit images.|
|_OneBitDownsampleThreshold_|Optional| **Long**|The threshold at or above which an image is down-sampled to the _OneBitDownsampleTarget_ level.|
|_From_|Optional| **Long**|The page number of the first page to export.|
|_To_|Optional| **Long**|The page number of the last page to export.|
|_Copies_|Optional| **Long**|The number of copies.|
|_Collate_|Optional| **Boolean**|Whether to collate the copies.|
|_PrintStyle_|Optional| **[PbPrintStyle](Publisher.PbPrintStyle.md)**|The style in which to print the exported file. Can be one of the **PbPrintStyle** constants. The default value depends on the value of the _Intent_ parameter.|
|_DocStructureTags_|Optional| **Boolean**|Whether to include document structure tags to improve document accessibility. The default is **True**.|
|_BitmapMissingFonts_|Optional| **Boolean**|Whether to include a bitmap of the text. Pass **True** for this parameter when font licenses do not permit a font to be embedded in the PDF file. If you pass **False**, the font is referenced, and the viewer's computer substitutes an appropriate font if the authored one is not available. Default value is **True**. |
|_UseISO19005\_1_|Optional| **Boolean**|Whether the resulting document is compliant with ISO 19005-1 (PDF/A). The default is **False**.|
|_ExternalExporter_|Optional| **Variant**|A pointer to an add-in that allows calls to an alternate implementation of code. You can add support for additional fixed formats by writing a Microsoft Office add-in that implements the **IMsoDocExporter** COM interface. For more information, see [Extend the fixed-format export feature in Word Automation Services](https://docs.microsoft.com/sharepoint/dev/general-development/extend-the-fixed-format-export-feature-in-word-automation-services).|

## Remarks

The **ExportAsFixedFormat** method is the equivalent of the **Publish As PDF or XPS** command on the **File** menu in the Publisher user interface.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **ExportAsFixedFormat** method to save the active publication as a .pdf file.

Before running this code, replace `pathandfilename.pdf` with a valid file name and the path to a folder on your computer where you have permission to save files.

```vb
Public Sub ExportAsFixedFormat_Example() 
 
 ThisDocument.ExportAsFixedFormat pbFixedFormatTypePDF, "pathandfilename.pdf" 
 
End Sub
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]