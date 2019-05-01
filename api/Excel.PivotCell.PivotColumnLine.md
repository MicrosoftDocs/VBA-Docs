---
title: PivotCell.PivotColumnLine property (Excel)
keywords: vbaxl10.chm692084
f1_keywords:
- vbaxl10.chm692084
ms.prod: excel
api_name:
- Excel.PivotCell.PivotColumnLine
ms.assetid: 99d8e14e-28b5-4c0c-2f92-402fbb5c2ea8
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotCell.PivotColumnLine property (Excel)

Returns the  **PivotLine** on a column for a specific **PivotCell** object. Read-only **PivotLine**.


## Syntax

_expression_. `PivotColumnLine`

_expression_ A variable that represents a **[PivotCell](Excel.PivotCell.md)** object.


## Remarks

If the PivotCell is on rows, the  **PivotColumnLine** property returns a run-time error.

If the PivotCell is on columns, the  **PivotColumnLine** property returns the column **PivotLine** object.

If the PivotCell is in the data area, the  **PivotColumnLine** property returns the corresponding column **PivotLine** object.


## See also


[PivotCell Object](Excel.PivotCell.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]