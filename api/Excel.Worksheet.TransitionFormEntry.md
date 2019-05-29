---
title: Worksheet.TransitionFormEntry property (Excel)
keywords: vbaxl10.chm175132
f1_keywords:
- vbaxl10.chm175132
ms.prod: excel
api_name:
- Excel.Worksheet.TransitionFormEntry
ms.assetid: ec17c4db-d94e-2fd9-39eb-7c1e0ea40a49
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.TransitionFormEntry property (Excel)

**True** if Microsoft Excel uses Lotus 1-2-3 formula entry rules for the worksheet. Read/write **Boolean**.


## Syntax

_expression_.**TransitionFormEntry**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Example

This example causes Excel to use Lotus 1-2-3 formula entry rules for Sheet1.

```vb
Worksheets("Sheet1").TransitionFormEntry = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]