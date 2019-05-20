---
title: Window.GridlineColor property (Excel)
keywords: vbaxl10.chm356093
f1_keywords:
- vbaxl10.chm356093
ms.prod: excel
api_name:
- Excel.Window.GridlineColor
ms.assetid: d2d35a5c-cc5c-4547-a22d-78fe2ef11073
ms.date: 05/21/2019
localization_priority: Normal
---


# Window.GridlineColor property (Excel)

Returns or sets the gridline color as an RGB value. Read/write **Long**.


## Syntax

_expression_.**GridlineColor**

_expression_ A variable that represents a **[Window](Excel.Window.md)** object.


## Example

This example sets the gridline color in the active window in Book1.xls to red.

```vb
Workbooks("BOOK1.XLS").Worksheets("Sheet1").Activate 
ActiveWindow.GridlineColor = RGB(255,0,0)
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]