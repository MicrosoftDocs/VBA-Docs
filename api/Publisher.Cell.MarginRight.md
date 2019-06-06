---
title: Cell.MarginRight property (Publisher)
keywords: vbapb10.chm5111828
f1_keywords:
- vbapb10.chm5111828
ms.prod: publisher
api_name:
- Publisher.Cell.MarginRight
ms.assetid: d297222e-7fc1-9225-e098-1a85d7734d77
ms.date: 06/06/2019
localization_priority: Normal
---


# Cell.MarginRight property (Publisher)

Returns or sets a **Variant** that represents the amount of space (in [points](../language/glossary/vbe-glossary.md#point)) between the text and the right edge of a cell, text frame, or page. Read/write.


## Syntax

_expression_.**MarginRight**

_expression_ A variable that represents a **[Cell](Publisher.Cell.md)** object.


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