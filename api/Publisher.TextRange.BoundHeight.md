---
title: TextRange.BoundHeight property (Publisher)
keywords: vbapb10.chm5308436
f1_keywords:
- vbapb10.chm5308436
ms.prod: publisher
api_name:
- Publisher.TextRange.BoundHeight
ms.assetid: 010d3de9-5838-fbf7-fb75-b80a06aafac8
ms.date: 06/15/2019
localization_priority: Normal
---


# TextRange.BoundHeight property (Publisher)

Returns a **Single** indicating the height, in [points](../language/glossary/vbe-glossary.md#point), of the bounding box for the specified text range. Read-only.


## Syntax

_expression_.**BoundHeight**

_expression_ A variable that represents a **[TextRange](Publisher.TextRange.md)** object.


## Return value

Single


## Example

The following example displays the position, width, and height of the bounding box surrounding the text in the first shape on page one of the active publication.

```vb
Dim rngText As TextRange 
Dim strMessage As String 
 
Set rngText = ActiveDocument.Pages(1) _ 
 .Shapes(1).TextFrame.TextRange 
 
With rngText 
 strMessage = "Text frame information" & vbCrLf _ 
 & " Distance from left edge of page: " _ 
 & .BoundLeft & " points" & vbCrLf _ 
 & " Distance from top edge of page: " _ 
 & .BoundTop & " points" & vbCrLf _ 
 & " Width: " & .BoundWidth & " points" & vbCrLf _ 
 & " Height: " & .BoundHeight & " points" 
End With 
 
MsgBox strMessage
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]