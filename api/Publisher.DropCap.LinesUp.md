---
title: DropCap.LinesUp property (Publisher)
keywords: vbapb10.chm5505031
f1_keywords:
- vbapb10.chm5505031
ms.prod: publisher
api_name:
- Publisher.DropCap.LinesUp
ms.assetid: 97bf3fc1-2203-d916-0c2d-352260c279fe
ms.date: 06/07/2019
localization_priority: Normal
---


# DropCap.LinesUp property (Publisher)

Returns or sets a **Long** that represents the number of lines that an initial dropped capital letter is raised above the line of text on which it exists. Read/write.


## Syntax

_expression_.**LinesUp**

_expression_ A variable that represents a **[DropCap](Publisher.DropCap.md)** object.


## Return value

Long


## Example

This example creates a custom dropped capital letter that is five lines high and raises it two lines above the line on which it exists.

```vb
Sub RaisedDropCap() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes _ 
 .AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=100, Top:=100, Width:=100, Height:=100) _ 
 .TextFrame.TextRange 
 For intCount = 1 To 10 
 .InsertAfter NewText:="This is a test. " 
 Next intCount 
 With .DropCap 
 .Size = 5 
 .LinesUp = 2 
 End With 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]