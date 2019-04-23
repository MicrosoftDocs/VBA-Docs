---
title: DeleteReplacement method (Excel Graph)
keywords: vbagr10.chm66683
f1_keywords:
- vbagr10.chm66683
ms.prod: excel
api_name:
- Excel.DeleteReplacement
ms.assetid: d82693f6-5275-2473-55e8-2b3cc156d702
ms.date: 04/09/2019
localization_priority: Normal
---


# DeleteReplacement method (Excel Graph)

Deletes an entry from the array of AutoCorrect replacements.

## Syntax

_expression_.**DeleteReplacement** (_What_)

_expression_ Required. An expression that returns an **[AutoCorrect](excel.autocorrect-graph-object.md)** object.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_What_ | Required |**String**|The text to be replaced, as it appears in the row to be deleted from the array of AutoCorrect replacements. If this string doesn't exist in the array of AutoCorrect replacements, this method fails.|

## Example

This example removes the word Temperature from the array of AutoCorrect replacements.

```vb
With myChart.Application.AutoCorrect 
 .DeleteReplacement "Temperature" 
End With
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]