---
title: Range.HasFormula property (Excel)
keywords: vbaxl10.chm144143
f1_keywords:
- vbaxl10.chm144143
ms.prod: excel
api_name:
- Excel.Range.HasFormula
ms.assetid: a18bea77-cee9-ae2d-7e97-90a4205e3b1f
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.HasFormula property (Excel)

**True** if all cells in the range contain formulas; **False** if none of the cells in the range contains a formula; **null** otherwise. Read-only **Variant**.


## Syntax

_expression_.**HasFormula**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Example

This example prompts the user to select a range on Sheet1. If every cell in the selected range contains a formula, the example displays a message.

```vb
Worksheets("Sheet1").Activate 
Set rr = Application.InputBox( _ 
 prompt:="Select a range on this worksheet", _ 
 Type:=8) 
If rr.HasFormula = True Then 
 MsgBox "Every cell in the selection contains a formula" 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
