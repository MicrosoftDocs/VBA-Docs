---
title: Range.SpecialCells method (Excel)
keywords: vbaxl10.chm144203
f1_keywords:
- vbaxl10.chm144203
ms.prod: excel
api_name:
- Excel.Range.SpecialCells
ms.assetid: 30c2035c-34e3-3b1a-f243-69a9fed97f3b
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.SpecialCells method (Excel)

Returns a **Range** object that represents all the cells that match the specified type and value.


## Syntax

_expression_.**SpecialCells** (_Type_, _Value_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[XlCellType](Excel.XlCellType.md)**|The cells to include.|
| _Value_|Optional| **Variant**|If _Type_ is either **xlCellTypeConstants** or **xlCellTypeFormulas**, this argument is used to determine which types of cells to include in the result. These values can be added together to return more than one type. The default is to select all constants or formulas, no matter what the type.|

## Return value

Range


## Remarks

Use the **[XlSpecialCellsValue](excel.xlspecialcellsvalue.md)** enumeration to specify cells with a particular type of value to include in the result.

## Example

This example selects the last cell in the used range of Sheet1.

```vb
Worksheets("Sheet1").Activate 
ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Activate
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
