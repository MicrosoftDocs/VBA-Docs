---
title: DataBar.BarFillType property (Excel)
keywords: vbaxl10.chm810091
f1_keywords:
- vbaxl10.chm810091
ms.prod: excel
api_name:
- Excel.DataBar.BarFillType
ms.assetid: c83fc8d3-63aa-4989-8099-74bcad7d6fce
ms.date: 04/23/2019
localization_priority: Normal
---


# DataBar.BarFillType property (Excel)

Returns or sets how a data bar is filled with color. Read/write.


## Syntax

_expression_.**BarFillType**

_expression_ A variable that represents a **[DataBar](Excel.DataBar.md)** object.


## Return value

**[XlDataBarFillType](Excel.XlDataBarFillType.md)**


## Remarks

The default value of the **BarFillType** property is **xlDataBarFillGradient**.


## Example

The following code example selects a range of cells, adds a data bar conditional formatting rule to that range, and then sets the data bar's fill color to solid.

```vb
Range("A1:A10").Select 
Range("A1:A10").Activate 
 
Set myDataBar = Selection.FormatConditions.AddDatabar 
myDataBar.BarFillType = xlDataBarFillSolid
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]