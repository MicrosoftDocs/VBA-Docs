---
title: Range.MergeCells property (Excel)
keywords: vbaxl10.chm144161
f1_keywords:
- vbaxl10.chm144161
ms.prod: excel
api_name:
- Excel.Range.MergeCells
ms.assetid: 42904357-5e55-1eb0-9b06-83b446fc6275
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.MergeCells property (Excel)

**True** if the range contains merged cells. Read/write **Variant**.


## Syntax

_expression_.**MergeCells**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

When you select a range that contains merged cells, the resulting selection may be different from the intended selection. Use the **[Address](Excel.Range.Address.md)** property to check the address of the selected range.


## Example

This example sets the value of the merged range that contains cell A3.

```vb
Set ma = Range("a3").MergeArea 
If Range("a3").MergeCells Then 
 ma.Cells(1, 1).Value = "42" 
End If
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
