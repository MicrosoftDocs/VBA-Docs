---
title: DisplayFormat.FormulaHidden property (Excel)
keywords: vbaxl10.chm893078
f1_keywords:
- vbaxl10.chm893078
ms.prod: excel
api_name:
- Excel.DisplayFormat.FormulaHidden
ms.assetid: 3db0fd6b-da1b-f19a-e859-a949b5f4d2b3
ms.date: 04/25/2019
localization_priority: Normal
---


# DisplayFormat.FormulaHidden property (Excel)

Returns a value that indicates if the formula of the associated **[Range](Excel.Range(object).md)** object is hidden when the worksheet is protected as it is displayed in the current user interface. Read-only.


## Syntax

_expression_.**FormulaHidden**

_expression_ A variable that represents a **[DisplayFormat](Excel.DisplayFormat.md)** object.


## Return value

Variant


## Remarks

Returns **True** if the formula is hidden when the worksheet is protected. 

Returns **Null** if the range contains some cells with **FormulaHidden** equal to **True** and some cells with **FormulaHidden** equal to **False**.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]