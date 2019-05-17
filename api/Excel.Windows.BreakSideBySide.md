---
title: Windows.BreakSideBySide method (Excel)
keywords: vbaxl10.chm354079
f1_keywords:
- vbaxl10.chm354079
ms.prod: excel
api_name:
- Excel.Windows.BreakSideBySide
ms.assetid: be32b6a4-5541-8c4b-ef24-cf34c9035f1c
ms.date: 05/18/2019
localization_priority: Normal
---


# Windows.BreakSideBySide method (Excel)

Ends side-by-side mode if two windows are in side-by-side mode. Returns a **Boolean** value that represents whether the method was successful.


## Syntax

_expression_.**BreakSideBySide**

_expression_ A variable that represents a **[Windows](Excel.Windows.md)** object.


## Return value

**Boolean**


## Example

The following example ends side-by-side mode.

```vb
Sub CloseSideBySide() 
 
 ActiveWorkbook.Windows.BreakSideBySide 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]