---
title: Window.Activate method (Excel)
keywords: vbaxl10.chm356073
f1_keywords:
- vbaxl10.chm356073
ms.prod: excel
api_name:
- Excel.Window.Activate
ms.assetid: 7e0fdc4e-6399-62a8-f706-1653eb9217a2
ms.date: 06/08/2017
localization_priority: Priority
---


# Window.Activate method (Excel)

Brings the window to the front of the z-order. 


## Syntax

_expression_. `Activate`

_expression_ A variable that represents a [Window](./Excel.Window.md) object.


## Return value

Variant


## Remarks

This won't run any Auto_Activate or Auto_Deactivate macros that might be attached to the workbook (use the  **[RunAutoMacros](Excel.Workbook.RunAutoMacros.md)** method to run those macros).


## See also


[Window Object](Excel.Window.md)

