---
title: Worksheet.PivotTableAfterValueChange event (Excel)
keywords: vbaxl10.chm502082
f1_keywords:
- vbaxl10.chm502082
ms.prod: excel
api_name:
- Excel.Worksheet.PivotTableAfterValueChange
ms.assetid: 097e1c1e-4df6-a0d1-de67-0e0752d2286a
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.PivotTableAfterValueChange event (Excel)

Occurs after a cell or range of cells inside a PivotTable are edited or recalculated (for cells that contain formulas).


## Syntax

_expression_.**PivotTableAfterValueChange** (_TargetPivotTable_, _TargetRange_)

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _TargetPivotTable_|Required| **[PivotTable](Excel.PivotTable.md)**|The PivotTable that contains the edited or recalculated cells.|
| _TargetRange_|Required| **[Range](Excel.Range(object).md)**|The range that contains all the edited or recalculated cells.|

## Return value

**Nothing**


## Remarks

The **PivotTableAfterValueChange** event does not occur under any conditions other than editing or recalculating cells. For example, it will not occur when the PivotTable is refreshed, sorted, filtered, or drilled down on, even though those operations move cells and potentially retrieve new values from the OLAP data source.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]