---
title: Range.Locked property (Excel)
keywords: vbaxl10.chm144157
f1_keywords:
- vbaxl10.chm144157
ms.prod: excel
api_name:
- Excel.Range.Locked
ms.assetid: 93c5f21d-6429-3287-0992-c810b9a429a8
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.Locked property (Excel)

Returns or sets a **Variant** value that indicates if the object is locked.


## Syntax

_expression_.**Locked**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

This property returns **True** if the object is locked, **False** if the object can be modified when the sheet is protected, or **Null** if the specified range contains both locked and unlocked cells.


## Example

This example unlocks cells A1:G37 on Sheet1 so that they can be modified when the sheet is protected.

```vb
Worksheets("Sheet1").Range("A1:G37").Locked = False 
Worksheets("Sheet1").Protect
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
