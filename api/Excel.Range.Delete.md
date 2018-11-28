---
title: Range.Delete method (Excel)
keywords: vbaxl10.chm144115
f1_keywords:
- vbaxl10.chm144115
ms.prod: excel
api_name:
- Excel.Range.Delete
ms.assetid: 7d890cc5-5b5b-35f9-2d97-e4fe48f244ee
ms.date: 06/08/2017
---


# Range.Delete method (Excel)

Deletes the object.


## Syntax

 _expression_. `Delete`( `_Shift_` )

 _expression_ A variable that represents a [Range](excel.range-graph-property.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Shift_|Optional| **Variant**|Used only with  **[Range](Excel.Range(object).md)** objects. Specifies how to shift cells to replace deleted cells. Can be one of the following **[xlDeleteShiftDirection](Excel.XlDeleteShiftDirection.md)** constants: **xlShiftToLeft** or **xlShiftUp**. If this argument is omitted, Microsoft Excel decides based on the shape of the range.|

## Return value

Variant


## See also


[Range Object](Excel.Range(object).md)

