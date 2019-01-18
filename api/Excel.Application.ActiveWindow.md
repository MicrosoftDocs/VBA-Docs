---
title: Application.ActiveWindow property (Excel)
keywords: vbaxl10.chm132079
f1_keywords:
- vbaxl10.chm132079
ms.prod: excel
api_name:
- Excel.Application.ActiveWindow
ms.assetid: 8f788ad0-ae4e-2442-593c-9440e37100de
ms.date: 06/08/2017
localization_priority: Priority
---


# Application.ActiveWindow property (Excel)

Returns a **[Window](Excel.Window.md)** object that represents the active Excel window (the window on top). Read-only. Returns **Nothing** if there are no windows open.

## Syntax

_expression_. `ActiveWindow`

_expression_ A variable that represents an [Application](Excel.Application(object).md) object.

## Example

This example displays the name (**Caption** property) of the active Excel window.

```vb
MsgBox "The name of the active window is " & ActiveWindow.Caption
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]