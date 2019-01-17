---
title: Application.Height property (Excel)
keywords: vbaxl10.chm133145
f1_keywords:
- vbaxl10.chm133145
ms.prod: excel
api_name:
- Excel.Application.Height
ms.assetid: 2842f4c9-93b6-64a8-2394-72b47cf0cc83
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Height property (Excel)

Returns or sets a  **Double** value that represents the height, in points, of the main application window.


## Syntax

_expression_. `Height`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Remarks

 If the window is minimized, this property is read-only and refers to the height of the icon. If the window is maximized, this property cannot be set. Use the **[WindowState](Excel.Window.WindowState.md)** property to determine the window state.


## See also


[Application Object](Excel.Application(object).md)

