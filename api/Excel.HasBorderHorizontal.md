---
title: HasBorderHorizontal property (Excel Graph)
keywords: vbagr10.chm67207
f1_keywords:
- vbagr10.chm67207
api_name:
- Excel.HasBorderHorizontal
ms.assetid: 9d5a86ea-73f1-a149-8fc9-ce104cdb41a3
ms.date: 04/11/2019
ms.localizationpriority: medium
---


# HasBorderHorizontal property (Excel Graph)

**True** if the chart data table has horizontal cell borders. Read/write **Boolean**.

## Syntax

_expression_.**HasBorderHorizontal**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example causes the chart data table to be displayed with an outline border and no cell borders.

```vb
With myChart 
 .HasDataTable = True 
 With .DataTable 
 .HasBorderHorizontal = False 
 .HasBorderVertical = False 
 .HasBorderOutline = True 
 End With 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]