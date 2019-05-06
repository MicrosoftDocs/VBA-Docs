---
title: PivotItem.DrillTo method (Excel)
keywords: vbaxl10.chm246094
f1_keywords:
- vbaxl10.chm246094
ms.prod: excel
api_name:
- Excel.PivotItem.DrillTo
ms.assetid: 627806c2-834f-d217-1439-1e17bedd15c0
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotItem.DrillTo method (Excel)

The **DrillTo** method supports drilling to a specified PivotField from a PivotItem.


## Syntax

_expression_.**DrillTo** (_PivotItemName_)

_expression_ A variable that represents a **[PivotItem](Excel.PivotItem.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PivotItemName_|Required| **String**|The name of the PivotItem to drill to.|

## Remarks

For OLAP data sources, the PivotField being drilled to has to be in the same hierarchy as the PivotItem being drilled or, if multiple attribute hierarchies are placed next to each other on rows or columns, the PivotField being drilled to has to be one of the attribute hierarchies that are next to each other; no user hierarchies can be placed in between the PivotField of the PivotItem being drilled and the PivotField being drilled to. If these conditions are not met, a run-time error is returned.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]