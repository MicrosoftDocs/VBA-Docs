---
title: Window.Width property (Excel)
keywords: vbaxl10.chm356123
f1_keywords:
- vbaxl10.chm356123
ms.prod: excel
api_name:
- Excel.Window.Width
ms.assetid: 5271dd4c-2e0f-cad1-fbe8-dda602202dc1
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.Width property (Excel)

Returns or sets a  **Double** value that represents the width, in [points](../language/glossary/vbe-glossary.md#point), of the window.


## Syntax

_expression_.**Width**

_expression_ A variable that represents a [Window](Excel.Window.md) object.


## Remarks

Use the  **[UsableWidth](Excel.Window.UsableWidth.md)** property to determine the maximum size for the window. You cannot set this property if the window is maximized or minimized. Use the **[WindowState](Excel.Window.WindowState.md)** property to determine the window state.


## See also


[Window Object](Excel.Window.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]