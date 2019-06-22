---
title: Application.ActiveWindow property (PowerPoint)
keywords: vbapp10.chm502004
f1_keywords:
- vbapp10.chm502004
ms.prod: powerpoint
api_name:
- PowerPoint.Application.ActiveWindow
ms.assetid: 762c1c6a-1f8a-f47a-7b75-006c745caee0
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ActiveWindow property (PowerPoint)

Returns a  **[DocumentWindow](PowerPoint.DocumentWindow.md)** object that represents the active document window. Read-only.


## Syntax

_expression_.**ActiveWindow**

_expression_ A variable that represents an **[Application](PowerPoint.Application.md)** object.


## Return value

DocumentWindow


## Example

This example minimizes the active window.


```vb
Application.ActiveWindow.WindowState = ppWindowMinimized
```


## See also


[Application Object](PowerPoint.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]