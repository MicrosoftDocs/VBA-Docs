---
title: Chart.Unprotect method (Excel)
keywords: vbaxl10.chm148095
f1_keywords:
- vbaxl10.chm148095
ms.prod: excel
api_name:
- Excel.Chart.Unprotect
ms.assetid: 59a367bd-037b-84aa-5b2f-d532614ed347
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.Unprotect method (Excel)

Removes protection from a sheet or workbook. This method has no effect if the sheet or workbook isn't protected.


## Syntax

_expression_.**Unprotect** (_Password_)

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Password_|Optional| **Variant**|A string that denotes the case-sensitive password to use to unprotect the chart. If the chart isn't protected with a password, this argument is ignored.|

## Remarks

If you forget the password, you cannot unprotect the sheet or workbook. It's a good idea to keep a list of your passwords and their corresponding document names in a safe place.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]