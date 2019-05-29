---
title: Worksheet.DisplayPageBreaks property (Excel)
keywords: vbaxl10.chm175138
f1_keywords:
- vbaxl10.chm175138
ms.prod: excel
api_name:
- Excel.Worksheet.DisplayPageBreaks
ms.assetid: 95152278-2618-f200-9933-b6574a49e256
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.DisplayPageBreaks property (Excel)

**True** if page breaks (both automatic and manual) on the specified worksheet are displayed. Read/write **Boolean**.


## Syntax

_expression_.**DisplayPageBreaks**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Remarks

You can't set this property if you don't have a printer installed.


## Example

This example causes Sheet1 to display page breaks.

```vb
Worksheets("Sheet1").DisplayPageBreaks = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]