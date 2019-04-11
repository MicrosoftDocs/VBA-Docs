---
title: HasBorderVertical property (Excel Graph)
keywords: vbagr10.chm5207458
f1_keywords:
- vbagr10.chm5207458
ms.prod: excel
api_name:
- Excel.HasBorderVertical
ms.assetid: ee6f449d-369c-1953-8540-b8baa4b281ab
ms.date: 04/11/2019
localization_priority: Normal
---


# HasBorderVertical property (Excel Graph)

**True** if the chart data table has vertical cell borders. Read/write **Boolean**.

## Syntax

_expression_.**HasBorderVertical**

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