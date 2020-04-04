---
title: Shapes.AddMediaObject method (PowerPoint)
keywords: vbapp10.chm543025
f1_keywords:
- vbapp10.chm543025
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.AddMediaObject
ms.assetid: 7e2ab704-7fd4-86d7-3f61-8d69c13b5685
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddMediaObject method (PowerPoint)

Deprecated in PowerPoint 2013. See [Shapes.AddMediaObject2 method (PowerPoint)](PowerPoint.Shapes.AddMediaObject2.md). Creates a media object. Returns a **[Shape](PowerPoint.Shape.md)** object that represents the new media object.


## Syntax

_expression_. `AddMediaObject`( `_FileName_`, `_Left_`, `_Top_`, `_Width_`, `_Height_` )

_expression_ A variable that represents a **[Shapes](PowerPoint.Shapes.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**| Required **String**. The file from which the media object is to be created. If the path isn't specified, the current working folder is used.|
| _Left_|Optional|**Single**|The position (in points) of the upper-left corner of the media object's bounding box relative to the upper-left corner of the document.|
| _Top_|Optional|**Single**|The position (in points) of the upper-left corner of the media object's bounding box relative to the upper-left corner of the document.|
| _Width_|Optional|**Single**|The width of the media object's bounding box, in points.|
| _Height_|Optional|**Single**|The height of the media object's bounding box, in points.|

## Return value

Shape


## Example

This example adds the movie named "Clock.avi" to _myDocument_.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes.AddMediaObject FileName:="C:\Windows\clock.avi", _
    Left:=5, Top:=5, Width:=100, Height:=100
```


## See also


[Shapes Object](PowerPoint.Shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]