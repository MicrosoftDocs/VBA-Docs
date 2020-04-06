---
title: ShapeRange.PictureFormat property (PowerPoint)
keywords: vbapp10.chm548032
f1_keywords:
- vbapp10.chm548032
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.PictureFormat
ms.assetid: 5d51631d-1cd4-fbfc-9198-9d883b281821
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.PictureFormat property (PowerPoint)

Returns a **[PictureFormat](PowerPoint.PictureFormat.md)** object that contains picture formatting properties for the specified shape. Read-only.


## Syntax

_expression_.**PictureFormat**

_expression_ A variable that represents a **[ShapeRange](PowerPoint.ShapeRange.md)** object.


## Return value

PictureFormat


## Remarks

This property applies to  **Shape** or **ShapeRange** objects that represent pictures or OLE objects.


## Example

This example sets the brightness and contrast for shape one on _myDocument_. Shape one must be a picture or an OLE object.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).PictureFormat

    .Brightness = 0.3

    .Contrast = .75

End With
```


## See also


[ShapeRange Object](PowerPoint.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]