---
title: Range.Interior property (Excel)
keywords: vbaxl10.chm144150
f1_keywords:
- vbaxl10.chm144150
ms.prod: excel
api_name:
- Excel.Range.Interior
ms.assetid: 9599b0f7-9f52-627c-51e6-d8be8aeb9bbf
ms.date: 06/08/2017
localization_priority: Priority
---


# Range.Interior property (Excel)

Returns an  **[Interior](Excel.Interior(object).md)** object that represents the interior of the specified object.


## Syntax

_expression_. `Interior`

_expression_ A variable that represents a [Range](excel.range-graph-property.md) object.


## Example

This example sets the interior color for cell A1 on Sheet1 to cyan.


```vb
Sub SetColor() 
 
 Worksheets("Sheet1").Range("A1").Interior.ColorIndex = 8 ' Cyan 
 
End Sub
```


## See also


[Range Object](Excel.Range(object).md)

