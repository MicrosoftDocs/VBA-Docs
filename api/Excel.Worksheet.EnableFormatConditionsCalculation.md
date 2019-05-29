---
title: Worksheet.EnableFormatConditionsCalculation property (Excel)
keywords: vbaxl10.chm175161
f1_keywords:
- vbaxl10.chm175161
ms.prod: excel
api_name:
- Excel.Worksheet.EnableFormatConditionsCalculation
ms.assetid: f1f56d9f-3a0f-e3d4-f686-1a695a55604e
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.EnableFormatConditionsCalculation property (Excel)

Returns or sets if conditional formats occur automatically as needed. Read/write **Boolean**.


## Syntax

_expression_.**EnableFormatConditionsCalculation**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Remarks

When set to **True** (default), evaluation of conditional formats occurs automatically as needed. When set to **False**, conditional formats are not re-evaluated. Any previously applied conditional formatting will still be visible, but it will not update as cell values or AppliesTo ranges are changed. 

The purpose of this flag is to allow VBA programmers to configure a rule completely before evaluating it. This is particularly useful when condition is applied over a large range because performance can be slow in these cases.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]