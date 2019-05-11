---
title: Range.Insert method (Excel)
keywords: vbaxl10.chm144149
f1_keywords:
- vbaxl10.chm144149
ms.prod: excel
api_name:
- Excel.Range.Insert
ms.assetid: e612bbc8-3942-3349-f157-c0459805794a
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.Insert method (Excel)

Inserts a cell or a range of cells into the worksheet or macro sheet and shifts other cells away to make space.


## Syntax

_expression_.**Insert** (_Shift_, _CopyOrigin_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Shift_|Optional| **Variant**|Specifies which way to shift the cells. Can be one of the following **[XlInsertShiftDirection](Excel.XlInsertShiftDirection.md)** constants: **xlShiftToRight** or **xlShiftDown**. If this argument is omitted, Microsoft Excel decides based on the shape of the range.|
| _CopyOrigin_|Optional| **Variant**|The copy origin; that is, from where to copy the format for inserted cells. Can be one of the following **[XlInsertFormatOrigin](Excel.XlInsertFormatOrigin.md)** constants: **xlFormatFromLeftOrAbove** (default) or **xlFormatFromRightOrBelow**.|

## Return value

Variant

## Remarks

There is no value for _CopyOrigin_ that is equivalent to _Clear Formatting_ when inserting cells interactively in Excel. To achieve this, use the **[ClearFormats](Excel.Range.ClearFormats.md)** method.

```vb
With Range("B2:E5")
    .Insert xlShiftDown
    .ClearFormats
End With
```

## Example

This example inserts a row above row 2, copying the format from the row below (row 3) instead of from the header row.

```vb
Range("2:2").Insert CopyOrigin:=xlFormatFromRightOrBelow
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
