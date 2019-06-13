---
title: ShapeRange.SaveAsPicture method (Publisher)
keywords: vbapb10.chm2294050
f1_keywords:
- vbapb10.chm2294050
ms.prod: publisher
api_name:
- Publisher.ShapeRange.SaveAsPicture
ms.assetid: 0be9b741-8f11-a386-313b-231a3269883a
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.SaveAsPicture method (Publisher)

Saves a range of one or more shapes as a picture file.


## Syntax

_expression_.**SaveAsPicture** (_FileName_, _pbResolution_)

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_FileName_|Required| **String**|The path and file name of the new picture file that you want to create. The graphics format that the picture is saved in is determined by the file name extension (such as .jpg or .gif) that you specify.|
|_pbResolution_|Optional| **[PbPictureResolution](Publisher.PbPictureResolution.md)** |The resolution in which you want the picture to be saved. Can be one of the **PbPictureResolution** constants.|


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **SaveAsPicture** method to save all the shapes on the first page of the active publication as a .jpg picture file.

Before running this code, replace `filename.jpg` with a valid file name and the path to a folder on your computer where you have permission to save files.

```vb
Public Sub SaveAsPicture_Example() 
 
 Dim pubShapeRange As Publisher.ShapeRange 
 Set pubShapeRange = ThisDocument.Pages(1).Shapes.Range 
 
 pubShapeRange.SaveAsPicture "filename.jpg" 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]