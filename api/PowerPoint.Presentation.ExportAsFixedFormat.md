---
title: Presentation.ExportAsFixedFormat method (PowerPoint)
keywords: vbapp10.chm583096
f1_keywords:
- vbapp10.chm583096
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.ExportAsFixedFormat
ms.assetid: bad3c9cb-49d7-2fdd-5110-9c1ed6491b08
ms.date: 06/30/2022
ms.localizationpriority: medium
---

# Presentation.ExportAsFixedFormat method (PowerPoint)

Publishes a copy of a Microsoft PowerPoint presentation as a file in a fixed format, either PDF or XPS.

## Syntax

_expression_.**ExportAsFixedFormat** (_Path_, _FixedFormatType_, _Intent_, _FrameSlides_, _HandoutOrder_, _OutputType_, _PrintHiddenSlides_, _PrintRange_, _RangeType_, _SlideShowName_, _IncludeDocProperties_, _KeepIRMSettings_, _DocStructureTags_, _BitmapMissingFonts_, _UseISO19005\_1_, _ExternalExporter_)

_expression_ An expression that returns a **[Presentation](PowerPoint.Presentation.md)** object.

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
| _UseISO19005\_1_ |Optional|**Boolean**|Whether the resulting document is compliant with ISO 19005-1 (PDF/A). The default is **False**.|
| _ExternalExporter_|Optional|**Variant**|A pointer to an Office add-in that implements the **IMsoDocExporter** COM interface and allows calls to an alternate implementation of code. The default is a null pointer.|

## Remarks

The **ExportAsFixedFormat** method is the equivalent of the **Save As PDF or XPS** command on the **Office** menu in the PowerPoint user interface. The method creates a file that contains a static view of the active presentation.

The _FixedFormatType_ parameter value can be one of these **PpFixedFormatType** constants.

|Constant|Value|Description|
|:-----|:-----|:-----|
|**ppFixedFormatTypePDF**|2|Export to PDF format.|
|**ppFixedFormatTypeXPS**|1|Export to XPS format.|

<br/>

The _Intent_ parameter value can be one of these **PpFixedFormatIntent** constants.

|Constant|Description|
|:-----|:-----|
|**ppFixedFormatIntentPrint**|Intended to be published online and printed.|
|**ppFixedFormatIntentScreen**|The default. Intended to be published only online.|

<br/>

The _FrameSlides_ parameter value can be one of these **MsoTriState** constants.

|Constant|Description|
|:-----|:-----|
|**msoFalse**|The default. Does not frame exported slides.|
|**msoTrue**|Frames exported slides.|

<br/>

The _HandoutOrder_ parameter value can be one of these **PpPrintHandoutOrder** constants.

|Constant|Description|
|:-----|:-----|
|**ppPrintHandoutHorizontalFirst**|Prints handouts with consecutive slides displayed horizontally first (in horizontal rows).|
|**ppPrintHandoutVerticalFirst**|The default. Prints handouts with consecutive slides displayed vertically first (in vertical columns).|

<br/>

The _OutputType_ parameter value can be a combination of one or more of these **PpPrintOutputType** constants.

|Constant|Description|
|:-----|:-----|
|**ppPrintOutputBuildSlides**||
|**ppPrintOutputFourSlideHandouts**|Prints four slides per handout page.|
|**ppPrintOutputNineSlideHandouts**|Prints nine slides per handout page.|
|**ppPrintOutputNotesPages**|Prints notes pages.|
|**ppPrintOutputOneSlideHandouts**|Prints one slide per handout page.|
|**ppPrintOutputOutline**|Prints outline view.|
|**ppPrintOutputSixSlideHandouts**|Prints six slides per handout page.|
|**ppPrintOutputSlides**|Prints all slides in the presentation. The default.|
|**ppPrintOutputThreeSlideHandouts**|Prints three slides per handout page.|
|**ppPrintOutputTwoSlideHandouts**|Prints two slides per handout page.|

<br/>

The _PrintHiddenSlides_ parameter value can be one of these **MsoTriState** constants.

|Constant|Description|
|:-----|:-----|
|**msoFalse**|The default. Does not print hidden slides.|
|**msoTrue**|Prints hidden slides.|

<br/>

The _RangeType_ parameter value can be one of these **PpPrintRangeType** constants.

|Constant|Description|
|:-----|:-----|
|**ppPrintAll**|The default. Exports all slides.|
|**ppPrintCurrent**|Exports only the current slide.|
|**ppPrintNamedSlideShow**|Exports the named (custom) slide show specified in _SlideShowName_.|
|**ppPrintSelection**|Exports selected slides.|
|**ppPrintSlideRange**|Exports the specified slide range.|

Set _BitmapMissingFonts_ to **True** when font licensing does not permit you to embed a font in the PDF file. If you set this parameter to **False**, the font is referenced, and the viewer's computer substitutes an appropriate font if the authored one is not available.

The _KeepIRMSettings_ parameter behaves specially for PDF. It controls the retention of both labels and encryption to the output file. For more information, see [Manage sensitivity labels in Office apps](/microsoft-365/compliance/sensitivity-labels-office-apps?view=o365-worldwide#pdf-support&preserve-view=true).

Due to the interaction of partner add-ins creating PDFs in Office with encryption, Office will default the _KeepIRMSettings_ flag to **FALSE** until _second RMID_ releases. 

## Example

The following example shows how to use the **ExportAsFixedFormat** method to export the active presentation as a .pdf file named _test.pdf_ to the user's Documents folder.

```vb
Public Sub ExportAsFixedFormat_Example() 
 
       ActivePresentation.ExportAsFixedFormat "C:\Users\username \Documents\test.pdf", ppFixedFormatTypePDF, ppFixedFormatIntentScreen, msoCTrue, ppPrintHandoutHorizontalFirst, ppPrintOutputBuildSlides, msoFalse, , , , False, False, False, False, False 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
