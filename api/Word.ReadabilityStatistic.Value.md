---
title: ReadabilityStatistic.Value property (Word)
keywords: vbawd10.chm162463745
f1_keywords:
- vbawd10.chm162463745
ms.prod: word
api_name:
- Word.ReadabilityStatistic.Value
ms.assetid: 58f31b9b-00d9-dd15-da7d-0266f0b6bdc5
ms.date: 06/08/2017
localization_priority: Normal
---


# ReadabilityStatistic.Value property (Word)

Returns the value of the grammar statistic. Read-only  **Long**.


## Syntax

_expression_.**Value**

_expression_ Required. A variable that represents a '[ReadabilityStatistic](Word.ReadabilityStatistic.md)' object.


## Example

This example checks the grammar in the active document and then displays the Flesch reading-ease index.


```vb
ActiveDocument.CheckGrammar 
MsgBox ActiveDocument.ReadabilityStatistics( _ 
 "Flesch Reading Ease").Value
```


## See also


[ReadabilityStatistic Object](Word.ReadabilityStatistic.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]