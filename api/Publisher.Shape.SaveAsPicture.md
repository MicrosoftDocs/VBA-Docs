---
title: Shape.SaveAsPicture method (Publisher)
keywords: vbapb10.chm2228375
f1_keywords:
- vbapb10.chm2228375
ms.prod: publisher
api_name:
- Publisher.Shape.SaveAsPicture
ms.assetid: 2cc18a83-b947-ca8c-eab4-71a03b79b82b
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.SaveAsPicture method (Publisher)

Saves a single shape as a picture file.


## Syntax

_expression_.**SaveAsPicture** (_FileName_, _pbResolution_)

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_FileName_|Required| **String**|The path and file name of the new picture file that you want to create. The graphics format that the picture is saved in is determined by the file name extension (such as .jpg or .gif) that you specify.|
|_pbResolution_|Optional| **[PbPictureResolution](Publisher.PbPictureResolution.md)** |The resolution in which you want the picture to be saved. Can be one of the **PbPictureResolution** constants. |


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **SaveAsPicture** method to save the first shape in the shapes collection on the first page of the active publication as a .jpg picture file.

Before running this code, replace `filename.jpg` with a valid file name and the path to a folder on your computer where you have permission to save files.

```vb
Public Sub SaveAsPicture_Example() 
 
 ThisDocument.Pages(1).Shapes(1).SaveAsPicture "filename.jpg" 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
