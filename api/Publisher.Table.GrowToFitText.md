---
title: Table.GrowToFitText property (Publisher)
keywords: vbapb10.chm4784132
f1_keywords:
- vbapb10.chm4784132
ms.prod: publisher
api_name:
- Publisher.Table.GrowToFitText
ms.assetid: d8822df7-a252-a5bb-be26-83df8ec5eb94
ms.date: 06/14/2019
localization_priority: Normal
---


# Table.GrowToFitText property (Publisher)

**True** for cells in a table to increase vertically to fit text. Read/write.


## Syntax

_expression_.**GrowToFitText**

_expression_ A variable that represents a **[Table](Publisher.Table.md)** object.


## Return value

Boolean


## Example

This example sets each row of the specified table to 12 points, and the row height doesn't increase as text is added to the cells in the rows.

```vb
Sub DontEnlargeTableCells() 
 Dim rowTable As Row 
 With ActiveDocument.Pages(1).Shapes(1).Table 
 .GrowToFitText = False 
 For Each rowTable In .Rows 
 rowTable.Height = 12 
 Next 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]