---
title: TextFrame.HorizontalAnchor property (Word)
keywords: vbawd10.chm162665364
f1_keywords:
- vbawd10.chm162665364
ms.prod: word
api_name:
- Word.TextFrame.HorizontalAnchor
ms.assetid: 6e78d938-343c-304c-2a40-ccf747c4f15d
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame.HorizontalAnchor property (Word)

Returns or sets the horizontal alignment of text in a text frame. Read/write  **[MsoHorizontalAnchor](Office.MsoHorizontalAnchor.md)**.


## Syntax

_expression_.**HorizontalAnchor**

_expression_ A variable that represents a **[TextFrame](Word.TextFrame.md)** object.


## Example

The following code example shows how to set the alignment for the first shape in the active document to top center.


```vb
Public Sub HorizontalAnchor_Example() 
 
 With ActiveDocument.Shapes(1) 
 .TextFrame.HorizontalAnchor = msoAnchorCenter 
 .TextFrame.VerticalAnchor = msoAnchorTop 
 End With 
 
End Sub
```


## See also


[TextFrame Object](Word.TextFrame.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]