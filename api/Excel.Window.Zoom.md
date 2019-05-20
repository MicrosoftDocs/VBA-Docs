---
title: Window.Zoom property (Excel)
keywords: vbaxl10.chm356126
f1_keywords:
- vbaxl10.chm356126
ms.prod: excel
api_name:
- Excel.Window.Zoom
ms.assetid: 82e6ac47-7054-52a9-383e-80be278dab0f
ms.date: 05/21/2019
localization_priority: Normal
---


# Window.Zoom property (Excel)

Returns or sets a **Variant** value that represents the display size of the window, as a percentage (100 equals normal size, 200 equals double size, and so on).


## Syntax

_expression_.**Zoom**

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Remarks

You can also set this property to **True** to make the window size fit the current selection.

This function affects only the sheet that's currently active in the window. To use this property on other sheets, you must first activate them.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]