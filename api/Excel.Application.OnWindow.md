---
title: Application.OnWindow property (Excel)
keywords: vbaxl10.chm133186
f1_keywords:
- vbaxl10.chm133186
ms.prod: excel
api_name:
- Excel.Application.OnWindow
ms.assetid: 73ae5d34-66e6-3c1e-07f8-08850d13a4f5
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.OnWindow property (Excel)

Returns or sets the name of the procedure that's run whenever you activate a window. Read/write  **String**.


## Syntax

_expression_. `OnWindow`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Remarks

The procedure specified by this property isn't run when other procedures switch to the window or when a command to switch to a window is received through a DDE channel. Instead, the procedure responds to the user's actions, such as clicking a window with the mouse.

If a worksheet or macro sheet has an Auto_Activate or Auto_Deactivate macro defined for it, those macros will be run after the procedure specified by the  **[OnWindow](Excel.Application.OnWindow.md)** property.


## See also


[Application Object](Excel.Application(object).md)

