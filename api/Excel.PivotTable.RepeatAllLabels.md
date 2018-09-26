---
title: PivotTable.RepeatAllLabels Method (Excel)
keywords: vbaxl10.chm235195
f1_keywords:
- vbaxl10.chm235195
ms.prod: excel
api_name:
- Excel.PivotTable.RepeatAllLabels
ms.assetid: 4ca1a7fa-4db6-20da-e37b-37445fee30cf
ms.date: 06/08/2017
---


# PivotTable.RepeatAllLabels Method (Excel)

Specifies whether to repeat item labels for all PivotFields in the specified PivotTable.


## Syntax

 _expression_. `RepeatAllLabels`( `_Repeat_` )

 _expression_ A variable that represents a '[PivotTable](Excel.PivotTable.md)' object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Repeat_|Required| **[XlPivotFieldRepeatLabels](Excel.XlPivotFieldRepeatLabels.md)**||

### Return value

Nothing


## Remarks

Using the  **RepeatAllLabels** method corresponds to the **Repeat All Item Labels** and **Do Not Repeat Item Labels** commands on the **Report Layout** drop-down list of the **PivotTable Tools Design** tab.

To specify whether to repeat item labels for a single PivotField, use the  **[RepeatLabels](Excel.PivotField.RepeatLabels.md)** property.


## See also


[PivotTable Object](Excel.PivotTable.md)

