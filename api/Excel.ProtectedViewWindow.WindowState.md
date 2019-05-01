---
title: ProtectedViewWindow.WindowState property (Excel)
keywords: vbaxl10.chm914083
f1_keywords:
- vbaxl10.chm914083
ms.prod: excel
api_name:
- Excel.ProtectedViewWindow.WindowState
ms.assetid: 9fd61fb6-1804-7eba-d1e3-a42b8500a52e
ms.date: 06/08/2017
localization_priority: Normal
---


# ProtectedViewWindow.WindowState property (Excel)

Returns or sets the state of the specified Protected View window. Read/write


## Syntax

_expression_.**WindowState**

_expression_ A variable that represents a **[ProtectedViewWindow](Excel.ProtectedViewWindow.md)** object.


## Return value

 **[XlProtectedViewWindowState](Excel.XlProtectedViewWindowState.md)**


## Example

The following code example maximizes the active Protected View window.


```vb
ActiveProtectedViewWindow.WindowState = xlProtectedViewWindowMaximized 
```


## See also


[ProtectedViewWindow Object](Excel.ProtectedViewWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]