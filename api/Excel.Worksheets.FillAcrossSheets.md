---
title: Worksheets.FillAcrossSheets method (Excel)
keywords: vbaxl10.chm470077
f1_keywords:
- vbaxl10.chm470077
ms.prod: excel
api_name:
- Excel.Worksheets.FillAcrossSheets
ms.assetid: c006cee2-67a1-2f24-3061-a2eb32ee9ecf
ms.date: 05/18/2019
localization_priority: Normal
---


# Worksheets.FillAcrossSheets method (Excel)

Copies a range to the same area on all other worksheets in a collection.


## Syntax

_expression_.**FillAcrossSheets** (_Range_, _Type_)

_expression_ A variable that represents a **[Worksheets](Excel.Worksheets.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range**|The range to fill on all the worksheets in the collection. The range must be from a worksheet within the collection.|
| _Type_|Optional| **[XlFillWith](Excel.XlFillWith.md)**|Specifies how to copy the range.|


## Example

This example fills the range A1:C5 on Sheet1, Sheet5, and Sheet7 with the contents of the same range on Sheet1.

```vb
x = Array("Sheet1", "Sheet5", "Sheet7") 
Sheets(x).FillAcrossSheets _ 
 Worksheets("Sheet1").Range("A1:C5")
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]