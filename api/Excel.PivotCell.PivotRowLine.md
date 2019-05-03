---
title: PivotCell.PivotRowLine property (Excel)
keywords: vbaxl10.chm692083
f1_keywords:
- vbaxl10.chm692083
ms.prod: excel
api_name:
- Excel.PivotCell.PivotRowLine
ms.assetid: e7e1ed02-b401-15b1-8548-fbdeb84796fc
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotCell.PivotRowLine property (Excel)

Returns the **[PivotLine](excel.pivotline.md)** object on a row for a specific **PivotCell** object. Read-only **PivotLine**.


## Syntax

_expression_.**PivotRowLine**

_expression_ A variable that represents a **[PivotCell](Excel.PivotCell.md)** object.


## Remarks

If the PivotCell is on rows, the **PivotRowLine** property returns the row's **PivotLine** object.

If the PivotCell is on columns, the **PivotRowLine** property returns a run-time error.

If the PivotCell is in the data area, the **PivotRowLine** property returns the corresponding row's **PivotLine** object.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]