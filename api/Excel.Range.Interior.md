---
title: Range.Interior property (Excel)
keywords: vbaxl10.chm144150
f1_keywords:
- vbaxl10.chm144150
ms.prod: excel
api_name:
- Excel.Range.Interior
ms.assetid: 9599b0f7-9f52-627c-51e6-d8be8aeb9bbf
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.Interior property (Excel)

Returns an **[Interior](Excel.Interior(object).md)** object that represents the interior of the specified object.


## Syntax

_expression_.**Interior**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Example

This example sets the interior color for cell A1 on Sheet1 to cyan.

```vb
Sub SetColor() 
 
 Worksheets("Sheet1").Range("A1").Interior.ColorIndex = 8 ' Cyan 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
