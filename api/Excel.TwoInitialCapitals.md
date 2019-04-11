---
title: TwoInitialCapitals property (Excel Graph)
keywords: vbagr10.chm5208088
f1_keywords:
- vbagr10.chm5208088
ms.prod: excel
api_name:
- Excel.TwoInitialCapitals
ms.assetid: cf6931c7-11ee-77b0-feb2-e047f7cb58e6
ms.date: 04/12/2019
localization_priority: Normal
---


# TwoInitialCapitals property (Excel Graph)

**True** if words that begin with two capital letters are corrected automatically. Read/write **Boolean**.

## Syntax

_expression_.**TwoInitialCapitals**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example sets Graph to automatically correct words that begin with two capital letters.

```vb
With myChart.Application.AutoCorrect 
 .TwoInitialCapitals = True 
 .ReplaceText = True 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]