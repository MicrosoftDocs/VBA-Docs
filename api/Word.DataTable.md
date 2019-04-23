---
title: DataTable object (Word)
keywords: vbawd10.chm708
f1_keywords:
- vbawd10.chm708
ms.prod: word
api_name:
- Word.DataTable
ms.assetid: 4e6094ea-3d83-6ec0-9788-9d22b884beb2
ms.date: 06/08/2017
localization_priority: Normal
---


# DataTable object (Word)

Represents a chart data table.


## Example

Use the  **[DataTable](Word.Chart.DataTable.md)** property to return a **DataTable** object. The following example adds a data table with an outline border to embedded chart one.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.HasDataTable = True 
 .Chart.DataTable.HasBorderOutline = True 
 End If 
End With
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]