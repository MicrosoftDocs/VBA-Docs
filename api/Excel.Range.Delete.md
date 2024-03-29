---
title: Range.Delete method (Excel)
keywords: vbaxl10.chm144115
f1_keywords:
- vbaxl10.chm144115
api_name:
- Excel.Range.Delete
ms.assetid: 7d890cc5-5b5b-35f9-2d97-e4fe48f244ee
ms.date: 05/10/2019
ms.localizationpriority: medium
---


# Range.Delete method (Excel)

Deletes the object.


## Syntax

_expression_.**Delete** (_Shift_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Shift_|Optional| **Variant**|Used only with **Range** objects. Specifies how to shift cells to replace deleted cells.<br/><br/>Can be one of the following **[XlDeleteShiftDirection](Excel.XlDeleteShiftDirection.md)** constants: **xlShiftToLeft** or **xlShiftUp**.<br/><br/>If this argument is omitted, Microsoft Excel decides based on the shape of the range.|

## Return value

Variant



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
