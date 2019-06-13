---
title: PictureFormat.TransparencyColor property (Publisher)
keywords: vbapb10.chm3604743
f1_keywords:
- vbapb10.chm3604743
ms.prod: publisher
api_name:
- Publisher.PictureFormat.TransparencyColor
ms.assetid: 908d2e21-3e2a-b75b-a82d-454686b7ecb8
ms.date: 06/13/2019
localization_priority: Normal
---


# PictureFormat.TransparencyColor property (Publisher)

Returns or sets an **[MsoColorType](office.msocolortype.md)** constant that represents the transparency color. Read/write.


## Syntax

_expression_.**TransparencyColor**

_expression_ A variable that represents a **[PictureFormat](Publisher.PictureFormat.md)** object.


## Return value

MsoColorType


## Example

This example creates a picture on the first page and sets the transparency color to black.

```vb
Sub SetTransparentColor() 
 With ActiveDocument.Pages(1).Shapes.AddPicture( _ 
 FileName:="C:\My Pictures\Sample.gif", LinkToFile:=msoFalse, _ 
 SaveWithDocument:=msoTrue, Left:=36, Top:=36) 
 .PictureFormat.TransparencyColor = RGB(Red:=255, Green:=255, Blue:=255) 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]