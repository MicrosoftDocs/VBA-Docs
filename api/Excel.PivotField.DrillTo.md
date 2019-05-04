---
title: PivotField.DrillTo method (Excel)
keywords: vbaxl10.chm240138
f1_keywords:
- vbaxl10.chm240138
ms.prod: excel
api_name:
- Excel.PivotField.DrillTo
ms.assetid: a00fe83a-136d-45a3-d3aa-f7ea4d434001
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotField.DrillTo method (Excel)

The **DrillTo** method supports drilling to a specified PivotField from another PivotField.


## Syntax

_expression_.**DrillTo** (_PivotFieldName_)

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PivotFieldName_|Required| **String**|The name of the **PivotField** to drill to.|

## Remarks

This operation can only be performed on fields that are actually on the PivotTable. Therefore, either a hierarchy containing the requested PivotField has to be in rows or columns of the PivotTable, or the attribute/relational field has to be in rows or columns of the PivotTable placed next to at least one other attribute/relational field. Also, the field being drilled to has to be in the same hierarchy or attribute chain. If these conditions are not met, a run-time error occurs.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]