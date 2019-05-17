---
title: Cell.Split method (Word)
keywords: vbawd10.chm156106957
f1_keywords:
- vbawd10.chm156106957
ms.prod: word
api_name:
- Word.Cell.Split
ms.assetid: c7eb0d00-ff7e-a737-2083-e16f46ead256
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.Split method (Word)

Splits a single table cell into multiple cells.


## Syntax

_expression_.**Split** (_NumRows_, _NumColumns_)

_expression_ Required. A variable that represents a '[Cell](Word.Cell.md)' object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NumRows_|Optional| **Variant**|The number of rows that the cell or group of cells is to be split into.|
| _NumColumns_|Optional| **Variant**|The number of columns that the cell or group of cells is to be split into.|

## Example

This example splits the first cell in the first table into two cells.


```vb
ActiveDocument.Tables(1).Cell(1, 1).Split NumColumns:=2
```


## See also


[Cell Object](Word.Cell.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]