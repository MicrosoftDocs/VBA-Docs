---
title: CellFormat.Locked property (Excel)
keywords: vbaxl10.chm676085
f1_keywords:
- vbaxl10.chm676085
ms.prod: excel
api_name:
- Excel.CellFormat.Locked
ms.assetid: 6cf62248-2ef4-ba2a-61da-427775e5414a
ms.date: 04/16/2019
localization_priority: Normal
---


# CellFormat.Locked property (Excel)

Returns or sets a **Variant** value that indicates if the object is locked.


## Syntax

_expression_.**Locked**

_expression_ A variable that represents a **[CellFormat](Excel.CellFormat.md)** object.


## Remarks

This property returns **True** if the object is locked, **False** if the object can be modified when the sheet is protected, or **Null** if the specified range contains both locked and unlocked cells.


## Example

This example unlocks cells A1:G37 on Sheet1 so that they can be modified when the sheet is protected.

```vb
Worksheets("Sheet1").Range("A1:G37").Locked = False 
Worksheets("Sheet1").Protect
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]