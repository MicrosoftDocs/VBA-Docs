---
title: TextFrame.MarginRight property (Publisher)
keywords: vbapb10.chm3866646
f1_keywords:
- vbapb10.chm3866646
ms.prod: publisher
api_name:
- Publisher.TextFrame.MarginRight
ms.assetid: bdbde217-6a51-7823-ac93-8bbffa583544
ms.date: 06/15/2019
localization_priority: Normal
---


# TextFrame.MarginRight property (Publisher)

Returns or sets a **Variant** that represents the amount of space (in [points](../language/glossary/vbe-glossary.md#point)) between the text and the right edge of a cell, text frame, or page. Read/write.


## Syntax

_expression_.**MarginRight**

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