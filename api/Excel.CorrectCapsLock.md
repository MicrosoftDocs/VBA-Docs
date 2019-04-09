---
title: CorrectCapsLock property (Excel Graph)
keywords: vbagr10.chm5207254
f1_keywords:
- vbagr10.chm5207254
ms.prod: excel
api_name:
- Excel.CorrectCapsLock
ms.assetid: eb092056-2ae5-7982-28bb-1a367a812a9b
ms.date: 04/10/2019
localization_priority: Normal
---


# CorrectCapsLock property (Excel Graph)

**True** if Graph automatically corrects accidental use of the CapsLock key. Read/write **Boolean**.

## Syntax

_expression_.**CorrectCapsLock**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example enables Graph to automatically correct accidental use of the CapsLock key.

```vb
myChart.Application.AutoCorrect.CorrectCapsLock = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]