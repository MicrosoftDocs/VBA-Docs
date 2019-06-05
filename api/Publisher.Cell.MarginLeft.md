---
title: Cell.MarginLeft property (Publisher)
keywords: vbapb10.chm5111827
f1_keywords:
- vbapb10.chm5111827
ms.prod: publisher
api_name:
- Publisher.Cell.MarginLeft
ms.assetid: 1b665a3b-6958-0548-ece1-9d3a7045eaac
ms.date: 06/06/2019
localization_priority: Normal
---


# Cell.MarginLeft property (Publisher)

Returns or sets a **Variant** that represents the amount of space (in [points](../language/glossary/vbe-glossary.md#point)) between the text and the left edge of a cell, text frame, or page. Read/write.


## Syntax

_expression_.**MarginLeft**

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