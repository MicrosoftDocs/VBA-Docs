---
title: Shape.Export method (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.Shape.Export
ms.assetid: b5905350-17d7-4d63-ad39-570d967862a0
ms.date: 01/12/2024
ms.localizationpriority: medium
---


# Shape.Duplicate method (PowerPoint)

Exports a shape, using the specified graphics filter, and saves the exported file under the specified file name.

## Syntax

_expression_.**Export**(__Parameters__)

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|__PathName__|Required|String|The name of the file to be exported and saved to disk. You can include a full path; if you don't, Microsoft PowerPoint creates a file in the current folder. Specifies how far the shadow offset is to be moved horizontally, in points. A positive value moves the shadow to the right; a negative value moves it to the left.|
|__Filter__|Required|PpShapeFormat|The graphics filter to use in the creation of the exported image file.|
|__ScaleWidth__|Optional|Long|The width of the image in points. Default is the slide width.|
|__ScaleHeight__|Optional|Long|The height of the image in points. Default is the slide height.|
|__ExportMode__|Optional|ppExportMode|The scaling method use in the creation of the exported image file. If not specified, the dimensions will be scaled relative to the size of the slide.|

## Enumerations

### PpShapeFormat enumeration (PowerPoint)

|Name|Value|Description|
|:-----|:-----|:-----|
|ppShapeFormatBMP|3|Bitmap|
|ppShapeFormatEMF|5|Enhanced Metafile|
|ppShapeFormatGIF|0|Static GIF|
|ppShapeFormatJPG|1|Compressed JPG|
|ppShapeFormatPNG|2|Lossless PNG|
|ppShapeFormatSVG|6|Scalable Vector Graphic|
|ppShapeFormatWMF|4|Windows Metafile|

### ExportMode enumeration (PowerPoint)

|Name|Value|Description|
|:-----|:-----|:-----|
|ppClipRelativeToSlide |2|Reserved for future use |
|ppRelativeToSlide |1|Scales the image relative to the dimensions of the slide |
|ppScaleToFit |3|Reserved for future use |
|ppScaleXY |4|Reserved for future use |

## Remarks

PowerPoint uses the specified graphics filter to save each individual shape. The names of the shapes exported and saved to disk are determined the PathName argument which should include the corresponding file extension for the chosen graphics filter.

## See also

[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]