---
title: DisplayFormat.Locked property (Excel)
keywords: vbaxl10.chm893082
f1_keywords:
- vbaxl10.chm893082
ms.prod: excel
api_name:
- Excel.DisplayFormat.Locked
ms.assetid: 32941867-c714-cfa1-ad16-c214e745580e
ms.date: 04/25/2019
localization_priority: Normal
---


# DisplayFormat.Locked property (Excel)

Returns a value that indicates if the associated **[Range](Excel.Range(object).md)** object is locked as it is displayed in the current user interface. Read-only.


## Syntax

_expression_.**Locked**

_expression_ A variable that represents a **[DisplayFormat](Excel.DisplayFormat.md)** object.


## Return value

Variant


## Remarks

Returns **True** if the range is locked, **False** if the range can be modified when the sheet is protected, or **Null** if the range contains both locked and unlocked cells.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]