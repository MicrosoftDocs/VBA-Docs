---
title: TextFrame.MarginTop property (Publisher)
keywords: vbapb10.chm3866645
f1_keywords:
- vbapb10.chm3866645
ms.prod: publisher
api_name:
- Publisher.TextFrame.MarginTop
ms.assetid: 9709eefe-0857-f228-aa56-780c4789a413
ms.date: 06/15/2019
localization_priority: Normal
---


# TextFrame.MarginTop property (Publisher)

Returns or sets a **Variant** that represents the amount of space (in [points](../language/glossary/vbe-glossary.md#point)) between the text and the top edge of a cell, text frame, or page. Read/write.


## Syntax

_expression_.**MarginTop**

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