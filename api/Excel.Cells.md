---
title: Cells property (Excel Graph)
keywords: vbagr10.chm65774
f1_keywords:
- vbagr10.chm65774
ms.prod: excel
api_name:
- Excel.Cells
ms.assetid: 43d4d8ba-ae6b-90b8-6f83-bbb75a7cbccb
ms.date: 04/10/2019
localization_priority: Normal
---


# Cells property (Excel Graph)

Returns a **Range** object that represents the cells in the specified range, as it applies to the **Range** object. Also, returns a **Range** object that represents all the cells on the datasheet (not just the cells that are currently in use), as it applies to the **[DataSheet](excel.datasheet-graph-object.md)** object. Read-only **Range** object.

## Syntax

_expression_.**Cells**

_expression_ Required. An expression that returns a **[Range](excel.range-graph-object.md)** object.


## Example

This example clears the formula in cell A1 on the datasheet. Note that on the datasheet, column A is the second column and row 1 is the second row.

```vb
myChart.Application.DataSheet.Cells(2,2).ClearContents
```

<br/>

This example loops through cells A1:I3 on the datasheet. If any of these cells contains a value less than 0.001, the example replaces that value with 0 (zero).

```vb
Set mySheet = myChart.Application.DataSheet 
For rwIndex = 2 to 4 
 For colIndex = 2 to 10 
 If mySheet.Cells(rwIndex, colIndex) < .001 Then 
 mySheet.Cells(rwIndex, colIndex).Value = 0 
 End If 
 Next colIndex 
Next rwIndex
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]