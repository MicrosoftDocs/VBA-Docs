---
title: Window.Activate method (Excel)
keywords: vbaxl10.chm356073
f1_keywords:
- vbaxl10.chm356073
api_name:
- Excel.Window.Activate
ms.assetid: 7e0fdc4e-6399-62a8-f706-1653eb9217a2
ms.date: 05/21/2019
ms.localizationpriority: medium
---


# Window.Activate method (Excel)

Brings the window to the front of the z-order. 


## Syntax

_expression_.**Activate**

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Return value

Variant


## Remarks

This won't run any Auto_Activate or Auto_Deactivate macros that might be attached to the workbook (use the **[RunAutoMacros](Excel.Workbook.RunAutoMacros.md)** method to run those macros).




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
