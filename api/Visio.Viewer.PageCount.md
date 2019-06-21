---
title: Viewer.PageCount property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.PageCount
ms.assetid: 3a7f90c0-6573-7ba5-414d-ede5b9c2fac6
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.PageCount property (Visio Viewer)

Gets the number of pages in the current document that is open in Microsoft Visio Viewer. Read-only.


## Syntax

_expression_.**PageCount**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

**Long**


## Example

The following code displays the number of pages in the current document in the Immediate window.

```vb
Debug.Print vsoViewer.PageCount
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]