---
title: Chart.Previous property (Excel)
keywords: vbaxl10.chm148086
f1_keywords:
- vbaxl10.chm148086
ms.prod: excel
api_name:
- Excel.Chart.Previous
ms.assetid: c0cf65c3-6e9f-7e04-9161-13ba118f23f1
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.Previous property (Excel)

Returns a **[Worksheet](Excel.Worksheet.md)** object that represents the previous sheet.


## Syntax

_expression_.**Previous**

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Remarks

If the object is a range, this property emulates pressing Shift+Tab; unlike the key combination, however, the property returns the previous cell without selecting it.

On a protected sheet, this property returns the previous unlocked cell. On an unprotected sheet, this property always returns the cell immediately to the left of the specified cell.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]