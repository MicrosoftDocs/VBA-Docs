---
title: Worksheet.UsedRange property (Excel)
keywords: vbaxl10.chm175134
f1_keywords:
- vbaxl10.chm175134
ms.prod: excel
api_name:
- Excel.Worksheet.UsedRange
ms.assetid: f004b93c-d785-de19-1fb4-bbe0b2e9b6cd
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.UsedRange property (Excel)

Returns a **[Range](Excel.Range(object).md)** object that represents the used range on the specified worksheet. Read-only.


## Syntax

_expression_.**UsedRange**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Example

This example selects the used range on Sheet1.

```vb
Worksheets("Sheet1").Activate 
ActiveSheet.UsedRange.Select
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
