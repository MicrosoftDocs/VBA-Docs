---
title: ChartObject.BottomRightCell Property (Excel)
keywords: vbaxl10.chm494074
f1_keywords:
- vbaxl10.chm494074
ms.prod: excel
api_name:
- Excel.ChartObject.BottomRightCell
ms.assetid: e437e7d9-b8fb-0a55-9741-1b11dea714b7
ms.date: 06/08/2017
---


# ChartObject.BottomRightCell Property (Excel)

Returns a  **[Range](Excel.Range(object).md)** object that represents the cell that lies under the lower-right corner of the object. Read-only.


## Syntax

 _expression_. `BottomRightCell`

 _expression_ A variable that represents a [ChartObject](Excel.ChartObject.md) object.


## Example

This example displays the address of the cell beneath the lower-right corner of embedded chart one on Sheet1.


```vb
MsgBox "The bottom right corner is over cell " & _ 
 Worksheets("Sheet1").ChartObjects(1).BottomRightCell.Address
```


## See also


[ChartObject Object](Excel.ChartObject.md)

