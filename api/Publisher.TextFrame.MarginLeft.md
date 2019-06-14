---
title: TextFrame.MarginLeft property (Publisher)
keywords: vbapb10.chm3866644
f1_keywords:
- vbapb10.chm3866644
ms.prod: publisher
api_name:
- Publisher.TextFrame.MarginLeft
ms.assetid: 4e784b9f-9467-5a14-c211-589e69c3b8bc
ms.date: 06/15/2019
localization_priority: Normal
---


# TextFrame.MarginLeft property (Publisher)

Returns or sets a **Variant** that represents the amount of space (in [points](../language/glossary/vbe-glossary.md#point)) between the text and the left edge of a cell, text frame, or page. Read/write.


## Syntax

_expression_.**MarginLeft**

_expression_ A variable that represents a **[TextFrame](Publisher.TextFrame.md)** object.


## Example

This example sets the margins of the active publication to two inches.

```vb
Sub SetPageMargins() 
 
 With ActiveDocument.LayoutGuides 
 .MarginTop = Application.InchesToPoints(Value:=2) 
 .MarginBottom = Application.InchesToPoints(Value:=2) 
 .MarginLeft = Application.InchesToPoints(Value:=2) 
 .MarginRight = Application.InchesToPoints(Value:=2) 
 End With 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]