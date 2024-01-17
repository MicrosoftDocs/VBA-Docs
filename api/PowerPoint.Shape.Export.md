---
title: Shape.Export method (PowerPoint)
api_name:
- PowerPoint.Shape.Export
ms.assetid: b5905350-17d7-4d63-ad39-570d967862a0
ms.date: 01/12/2024
ms.localizationpriority: medium
---


# Shape.Export method (PowerPoint)

Exports a shape, using the specified graphics filter, and saves the exported file under the specified file name.

## Syntax

_expression_.**Export**(_Parameters_)

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_PathName_|Required|String|The name of the file to be exported and saved to disk. You can include a full path; if you don't, Microsoft PowerPoint creates a file in the current folder. Specifies how far the shadow offset is to be moved horizontally, in points. A positive value moves the shadow to the right; a negative value moves it to the left.|
|_Filter_|Required|PpShapeFormat|The graphics filter to use in the creation of the exported image file.|
|_ScaleWidth_|Optional|Long|The width of the image in points. Default is the slide width.|
|_ScaleHeight_|Optional|Long|The height of the image in points. Default is the slide height.|
|_ExportMode_|Optional|ppExportMode|The scaling method use in the creation of the exported image file. If not specified, the dimensions will be scaled relative to the size of the slide.|

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

The _ScaleWidth_ and _ScaleHeight_ parameters are used to scale the exported image size relative to the dimensions of the slide. For example, if a plain 1" square shape is created on a slide, it will measure as 72x72 points in the Object Model. When exported without using any scale factor, the default scale of 1:1 is applied, and PowerPoint will use 96DPI to create a 96x96 pixel image. If a scale factor of 2x is used as shown in example 2 below, the exported image will be 192x192 pixels.

If the slide and/or shape is not fully downloaded, this method fails and an error occurs. For more information about the Partial Documents, see [Work with Partial Documents](/office/vba/powerpoint/how-to/work-with-partial-documents).  

When exporting from PowerPoint on macOS, the files must be created in either the Office sandbox folder or the PowerPoint sandbox folder:

`/Users/[username]/Library/Group Containers/UBF8T346G9.Office`
`/Users/[username]/Library/Containers/com.microsoft.Powerpoint/Data`

If you attempt to use a path outside of these sandbox locations, and automation error is raised.

## Availability

The SVG filter is available on Windows version 2302 and later.

The Export method is available on macOS on version 16.82 and later.

## Example

The following example exports all SVG shapes in the active presentation as SVG files to the user’s Pictures folder. The default **PpRelativeToSlide** value is used for the _ExportMode_ parameter which means that the exported image will be  

```vb

For Each oSld In ActivePresentation.Slides
    For Each oShp In oSld.Shapes
        If oShp.Type = msoGraphic Then
            FileName = oShp.Name & ".svg"
            PathToFolder = Environ("USERPROFILE") & "\Pictures\"
            oShp.Export PathToFolder & FileName, ppShapeFormatSVG
        End If
    Next
Next 

```

The following example uses the scale feature to export the selected object at a size relative to the slide. In this case, the slide is a standard 16:9 size which is 960x540 points. The exported image is created at twice the size of its size on the slide.

```vb

PathToFile = Environ("USERPROFILE") & "\Pictures\export.png"

With ActiveWindow.Selection.ShapeRange(1)
        .Export PathToFile, ppShapeFormatPNG, 1920, 1080, ppRelativeToSlide
End With

```

## See also

[Shape Object](PowerPoint.Shape.md)

[PageSetup.SlideHeight](powerpoint.pagesetup.slideheight.md)

[PageSetup.SlideWidth](powerpoint.pagesetup.slidewidth.md)

[Work with Partial Documents](/office/vba/powerpoint/how-to/work-with-partial-documents)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
