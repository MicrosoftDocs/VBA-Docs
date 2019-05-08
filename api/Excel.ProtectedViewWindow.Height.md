---
title: ProtectedViewWindow.Height property (Excel)
keywords: vbaxl10.chm914076
f1_keywords:
- vbaxl10.chm914076
ms.prod: excel
api_name:
- Excel.ProtectedViewWindow.Height
ms.assetid: 32d5baad-2c78-02ad-7814-f703889f8a36
ms.date: 05/09/2019
localization_priority: Normal
---


# ProtectedViewWindow.Height property (Excel)

Returns or sets a value that represents the height, in [points](../language/glossary/vbe-glossary.md#point), of the Protected View window. Read/write.


## Syntax

_expression_.**Height**

_expression_ A variable that represents a **[ProtectedViewWindow](Excel.ProtectedViewWindow.md)** object.


## Return value

**Double**


## Remarks

You cannot set this property if the Protected View window is maximized or minimized. Use the **[WindowState](Excel.ProtectedViewWindow.WindowState.md)** property to determine the window state.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]