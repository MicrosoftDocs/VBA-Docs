---
title: ReplaceText property (Excel Graph)
keywords: vbagr10.chm5207920
f1_keywords:
- vbagr10.chm5207920
ms.prod: excel
api_name:
- Excel.ReplaceText
ms.assetid: 930c453b-5363-3124-ec06-62359e41ee47
ms.date: 04/12/2019
localization_priority: Normal
---


# ReplaceText property (Excel Graph)

**True** if the text in the list of AutoCorrect replacements is replaced automatically. Read/write **Boolean**.

## Syntax

_expression_.**ReplaceText**

_expression_ Required. An expression that returns an **[AutoCorrect](excel.autocorrect-graph-object.md)** object.

## Example

This example turns off automatic text replacement for the chart.

```vb
With myChart.Application.AutoCorrect 
 .CapitalizeNamesOfDays = True 
 .ReplaceText = False 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]