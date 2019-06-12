---
title: Page.PageNumber property (Publisher)
keywords: vbapb10.chm393220
f1_keywords:
- vbapb10.chm393220
ms.prod: publisher
api_name:
- Publisher.Page.PageNumber
ms.assetid: 670e3f46-9cad-b85e-b627-3be8c7c4e577
ms.date: 06/11/2019
localization_priority: Normal
---


# Page.PageNumber property (Publisher)

Returns a **String** that represents the current page number. Read-only.


## Syntax

_expression_.**PageNumber**

_expression_ A variable that represents a **[Page](Publisher.Page.md)** object.


## Return value

String


## Example

This example creates a text box, gets the current page number, and inserts it with new text into the new shape.


```vb
Sub GetPageNumber() 
 Dim strPageNumber As String 
 With ActiveDocument.Pages(1) 
 strPageNumber = .PageNumber 
 .Shapes.AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=100, Top:=100, Width:=100, Height:=100) _ 
 .TextFrame.TextRange.InsertAfter NewText:="This is page " _ 
 & strPageNumber & " of " & .Parent.Count & "." 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]