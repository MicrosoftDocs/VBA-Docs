---
title: PivotTable.DrillTo method (Excel)
keywords: vbaxl10.chm235208
f1_keywords:
- vbaxl10.chm235208
ms.prod: excel
ms.assetid: 9f700cba-2cf5-4b13-707f-254148ddf73a
ms.date: 06/08/2017
localization_priority: Normal
---


# PivotTable.DrillTo method (Excel)

Enables you to drill to a location within an OLAP or PowerPivot based cube hierarchy.


## Syntax

_expression_.**DrillTo** (_PivotItem_, _CubeField_, _PivotLine_)

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PivotItem_|Required|PIVOTITEM|The member from which the drill operation is performed.|
| _CubeField_|Required|CUBEFIELD|The target hierarchy being drilled to.|
| _PivotLine_|Optional|**Variant**|Specifies the line in the PivotTable where the operation starting member resides. In cases where PivotLine is not specified, defaults to the top PivotLine where the member appears.|

## Return value

 **VOID**


## See also


[PivotTable Object](Excel.PivotTable.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]