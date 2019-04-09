---
title: CorrectSentenceCap property (Excel Graph)
keywords: vbagr10.chm67155
f1_keywords:
- vbagr10.chm67155
ms.prod: excel
api_name:
- Excel.CorrectSentenceCap
ms.assetid: f0f5920d-fb2e-3a06-35ca-0e67202df6db
ms.date: 04/10/2019
localization_priority: Normal
---


# CorrectSentenceCap property (Excel Graph)

**True** if Graph automatically corrects sentence (first word) capitalization. Read/write **Boolean**.

## Syntax

_expression_.**CorrectSentenceCap**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example enables Graph to automatically correct sentence capitalization.

```vb
myChart.Application.AutoCorrect.CorrectSentenceCap = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]