---
title: TextRange.Length property (Publisher)
keywords: vbapb10.chm5308432
f1_keywords:
- vbapb10.chm5308432
ms.prod: publisher
api_name:
- Publisher.TextRange.Length
ms.assetid: 003b4ad1-2c09-17c9-279b-b1cf2ebdb40a
ms.date: 06/15/2019
localization_priority: Normal
---


# TextRange.Length property (Publisher)

Returns a **Long** value indicating the length of the specified text range, in characters. Read-only.


## Syntax

_expression_.**Length**

_expression_ A variable that represents a **[TextRange](Publisher.TextRange.md)** object.


## Example

This example sets the font size of a text frame on page two to 48 points if the text frame contains more than five characters, or it sets the font size to 72 points if the text frame has five or fewer characters.

```vb
With ActiveDocument.Pages(2).Shapes(1) _ 
 .TextFrame.TextRange 
 If .Length > 5 Then 
 .Font.Size = 48 
 Else 
 .Font.Size = 72 
 End If 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]