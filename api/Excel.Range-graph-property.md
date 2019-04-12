---
title: Range property (Excel Graph)
keywords: vbagr10.chm65733
f1_keywords:
- vbagr10.chm65733
ms.prod: excel
api_name:
- Excel.Range
ms.assetid: 760f463d-3af3-515d-5da4-54f799fcfe0b
ms.date: 04/12/2019
localization_priority: Normal
---

# Range property (Excel Graph)

Returns a **[Range](excel.range-graph-object.md)** object that represents the specified cell or range of cells. Read-only **Range** object.

## Syntax

_expression_.**Range** (_Range1_, _Range2_)

_expression_ Required. An expression that returns a **[DataSheet](excel.datasheet-graph-object.md)** object.

## Parameters

|Name|Required/Optional|Description|
|:-----|:-----|:-----|
|_Range1_ |Required for a single cell. |The name of the specified range.<br/><br/>This must be an A1-style reference in the language the macro is written in.<br/><br/>It can include:<ul><li>The range operator (a colon).</li><li>The intersection operator (a space).</li><li>The union operator (a comma).</li></ul>It can also include dollar signs, but they're ignored.|
|_Range1_, _Range2_ |Required for a range of cells. |The cells in the upper-left and lower-right corners of the specified range.<br/><br/>Each argument can be:<ul><li>A **Range** object that contains a single cell (or an entire column or entire row).</li><li>A string that names a single cell in the language that the macro is written in.</li></ul>|

## Remarks

On the datasheet, the first column heading (starting on the left) is A, followed by B, C, D, and so on. The first row heading (starting at the top) is 1, followed by 2, 3, 4, and so on. Neither the leftmost column nor the top row has a heading. In other words, column A is actually the second column from the left; likewise, row 1 is the second row from the top. 

The leftmost column and the top row, which are commonly used for legend text or axis labels, are referred to as column 0 (zero) and row 0 (zero). Thus, the following example inserts the text Annual Sales in the top cell in column A (the second column).

```vb
myChart.Application.DataSheet.Range("A0").Value = "Annual Sales"
```

<br/>

The following example inserts the text District 1 in the leftmost cell in row 2 (the third row).

```vb
myChart.Application.DataSheet.Range("02").Value = "District 1"
```

## Example

This example sets the value of cell A1 on the datasheet to 3.14159.

```vb
myChart.DataSheet.Range("A1").Value = 3.14159
```

<br/>

This example loops on cells A1:C3 on the datasheet. If one of the cells has a value of less than 0.001, the example replaces that value with 0 (zero).

```vb
With myChart.Application.DataSheet 
 For Each c in .Range("A1:C3") 
 If c.Value < .001 Then 
 c.Value = 0 
 End If 
 Next c 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
