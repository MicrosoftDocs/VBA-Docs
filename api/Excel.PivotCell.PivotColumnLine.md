---
title: PivotCell.PivotColumnLine property (Excel)
keywords: vbaxl10.chm692084
f1_keywords:
- vbaxl10.chm692084
ms.prod: excel
api_name:
- Excel.PivotCell.PivotColumnLine
ms.assetid: 99d8e14e-28b5-4c0c-2f92-402fbb5c2ea8
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotCell.PivotColumnLine property (Excel)

Returns the **[PivotLine](excel.pivotline.md)** object on a column for a specific **PivotCell** object. Read-only **PivotLine**.


## Syntax

_expression_.**PivotColumnLine**

_expression_ A variable that represents a **[PivotCell](Excel.PivotCell.md)** object.


## Remarks

If the PivotCell is on rows, the **PivotColumnLine** property returns a run-time error.

If the PivotCell is on columns, the **PivotColumnLine** property returns the column's **PivotLine** object.

If the PivotCell is in the data area, the **PivotColumnLine** property returns the corresponding column's **PivotLine** object.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]