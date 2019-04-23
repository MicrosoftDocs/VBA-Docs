---
title: AutoCorrect property (Excel Graph)
keywords: vbagr10.chm5207061
f1_keywords:
- vbagr10.chm5207061
ms.prod: excel
api_name:
- Excel.AutoCorrect
ms.assetid: f05a4ff5-4245-ff2e-1082-f48e130d0741
ms.date: 04/09/2019
localization_priority: Normal
---


# AutoCorrect property (Excel Graph)

Returns an **AutoCorrect** object that represents the Graph AutoCorrect attributes. Read-only.

## Syntax

_expression_.**AutoCorrect**

_expression_ Required. An expression that returns an **[AutoCorrect](Excel.AutoCorrect-graph-object.md)** object.

## Example

This example substitutes the word Temp. for the word Temperature in the array of AutoCorrect replacements.

```vb
With myChart.Application.AutoCorrect 
 .AddReplacement "Temperature", "Temp." 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]