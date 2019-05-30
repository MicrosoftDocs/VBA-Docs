---
title: Worksheet.TransitionExpEval property (Excel)
keywords: vbaxl10.chm175073
f1_keywords:
- vbaxl10.chm175073
ms.prod: excel
api_name:
- Excel.Worksheet.TransitionExpEval
ms.assetid: a92d8efb-5f19-4b41-11b2-a20b3ad5bf1d
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.TransitionExpEval property (Excel)

**True** if Microsoft Excel uses Lotus 1-2-3 expression evaluation rules for the worksheet. Read/write **Boolean**.


## Syntax

_expression_.**TransitionExpEval**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Example

This example causes Excel to use Lotus 1-2-3 expression evaluation rules for Sheet1.

```vb
Worksheets("Sheet1").TransitionExpEval = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]