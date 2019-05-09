---
title: Range.End property (Excel)
keywords: vbaxl10.chm144121
f1_keywords:
- vbaxl10.chm144121
ms.prod: excel
api_name:
- Excel.Range.End
ms.assetid: d46d75c9-b152-e93d-82c3-f59f0e7f69da
ms.date: 05/10/2019
localization_priority: Priority
---


# Range.End property (Excel)

Returns a **Range** object that represents the cell at the end of the region that contains the source range. Equivalent to pressing END+UP ARROW, END+DOWN ARROW, END+LEFT ARROW, or END+RIGHT ARROW. Read-only **Range** object.


## Syntax

_expression_.**End** (_Direction_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Direction_|Required| **[XlDirection](Excel.XlDirection.md)**|The direction in which to move.|

## Example

This example selects the cell at the top of column B in the region that contains cell B4.

```vb
Range("B4").End(xlUp).Select
```

<br>

This example selects the cell at the end of row 4 in the region that contains cell B4.

```vb
Range("B4").End(xlToRight).Select
```

<br>

This example extends the selection from cell B4 to the last cell in row four that contains data.

```vb
Worksheets("Sheet1").Activate 
Range("B4", Range("B4").End(xlToRight)).Select
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
