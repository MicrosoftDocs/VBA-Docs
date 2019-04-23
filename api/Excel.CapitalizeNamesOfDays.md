---
title: CapitalizeNamesOfDays property (Excel Graph)
keywords: vbagr10.chm66686
f1_keywords:
- vbagr10.chm66686
ms.prod: excel
api_name:
- Excel.CapitalizeNamesOfDays
ms.assetid: dbac8451-a2ac-5e29-b6c9-afa9cfaec469
ms.date: 04/10/2019
localization_priority: Normal
---


# CapitalizeNamesOfDays property (Excel Graph)

**True** if the first letter of day names is capitalized automatically. Read/write **Boolean**.

## Syntax

_expression_.**CapitalizeNamesOfDays**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Example

This example sets Graph to capitalize the first letter of the names of days.

```vb
With myChart.Application.AutoCorrect 
 .CapitalizeNamesOfDays = True 
 .ReplaceText = True 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]