---
title: Chart.Next property (Excel)
keywords: vbaxl10.chm148081
f1_keywords:
- vbaxl10.chm148081
ms.prod: excel
api_name:
- Excel.Chart.Next
ms.assetid: a0e53eba-c9e9-7997-4765-90debeb8ae5d
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.Next property (Excel)

Returns a **[Worksheet](Excel.Worksheet.md)** object that represents the next sheet.


## Syntax

_expression_.**Next**

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Remarks

If the object is a range, this property emulates the Tab key, although the property returns the next cell without selecting it.

On a protected sheet, this property returns the next unlocked cell. On an unprotected sheet, this property always returns the cell immediately to the right of the specified cell.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]