---
title: Sheets.FillAcrossSheets method (Excel)
keywords: vbaxl10.chm152077
f1_keywords:
- vbaxl10.chm152077
ms.prod: excel
api_name:
- Excel.Sheets.FillAcrossSheets
ms.assetid: eee9b0a2-0727-dfc8-ea7b-d7c582466d5c
ms.date: 06/08/2017
localization_priority: Normal
---


# Sheets.FillAcrossSheets method (Excel)

Copies a range to the same area on all other worksheets in a collection.


## Syntax

_expression_. `FillAcrossSheets`( `_Range_` , `_Type_` )

_expression_ A variable that represents a [Sheets](./Excel.Sheets.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**|The range to fill on all the worksheets in the collection. The range must be from a worksheet within the collection.|
| _Type_|Optional| **[xlFillWith](Excel.XlFillWith.md)**|Specifies how to copy the range.|

## Example

This example fills the range A1:C5 on Sheet1, Sheet5, and Sheet7 with the contents of the same range on Sheet1.


```vb
x = Array("Sheet1", "Sheet5", "Sheet7") 
Sheets(x).FillAcrossSheets _ 
 Worksheets("Sheet1").Range("A1:C5")
```


## See also


[Sheets Object](Excel.Sheets.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]