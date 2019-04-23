---
title: Application property (Excel Graph)
keywords: vbagr10.chm3076941
f1_keywords:
- vbagr10.chm3076941
ms.prod: excel
api_name:
- Excel.Application
ms.assetid: df183c1c-8db3-e85c-c390-977cf54db7c5
ms.date: 04/09/2019
localization_priority: Normal
---


# Application property (Excel Graph)

Returns an **Application** object that represents the Excel Graph application. Read-only **Application** object.

## Syntax

_expression_.**Application**

_expression_ Required. An expression that returns an **[Application](excel.application-graph-object.md)** object.


## Example

This example substitutes the word Temp. for the word Temperature in the array of AutoCorrect replacements.

```vb
With myChart.Application.AutoCorrect 
 .AddReplacement "Temperature", "Temp." 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
