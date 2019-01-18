---
title: Application.ActiveProtectedViewWindow property (Excel)
keywords: vbaxl10.chm133331
f1_keywords:
- vbaxl10.chm133331
ms.prod: excel
api_name:
- Excel.Application.ActiveProtectedViewWindow
ms.assetid: 2202c3b4-8880-7a26-8a56-8f2d2e7b7343
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ActiveProtectedViewWindow property (Excel)

Returns a  **[ProtectedViewWindow](Excel.ProtectedViewWindow.md)** object that represents the active **Protected View** window (the window on top). Read-only. Returns **Nothing** if there are no **Protected View** windows open. Read-only


## Syntax

_expression_. `ActiveProtectedViewWindow`

_expression_ A variable that represents an '[Application](Excel.Application(object).md)' object.


## Example

The following code example displays the name (**Caption** property) of the active **Protected View** window.


```vb
MsgBox "The name of the active Protected View window is " & ActiveProtectedWindow.Caption
```


## See also


[Application Object](Excel.Application(object).md)

