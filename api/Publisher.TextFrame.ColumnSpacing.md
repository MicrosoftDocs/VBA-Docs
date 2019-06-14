---
title: TextFrame.ColumnSpacing property (Publisher)
keywords: vbapb10.chm3866633
f1_keywords:
- vbapb10.chm3866633
ms.prod: publisher
api_name:
- Publisher.TextFrame.ColumnSpacing
ms.assetid: 3b650d29-3716-e9b1-eaf0-92bdc0b77c5f
ms.date: 06/15/2019
localization_priority: Normal
---


# TextFrame.ColumnSpacing property (Publisher)

Returns or sets a **Variant** that represents the amount of space between text columns. Read/write.


## Syntax

_expression_.**ColumnSpacing**

_expression_ A variable that represents a **[TextFrame](Publisher.TextFrame.md)** object.


## Return value

Variant


## Remarks

Spacing measures from the end of the text to the end of the column and again from the beginning of the column to the beginning of the text. 

Thus, if you enter a **ColumnSpacing** amount of 0.5 inch, the total spacing between columns is one inch: 0.5 inch measuring from the end of the text to the end of the column in one column, plus 0.5 inch measuring from the beginning of the column to the beginning of the text in a neighboring column.


## Example

This example formats the first text box in the active publication with three columns and a total of 0.5 inch spacing between columns.

```vb
Sub SetColumnsAndSpacing() 
 With ActiveDocument.Pages(1).Shapes(1).TextFrame 
 .Columns = 3 
 .ColumnSpacing = InchesToPoints(0.25) 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]