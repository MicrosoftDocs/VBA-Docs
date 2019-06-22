---
title: Viewer.PageVisible property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.PageVisible
ms.assetid: 7af34d35-b83d-931a-7116-fef8dab42f22
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.PageVisible property (Visio Viewer)

Gets or sets a value that indicates whether the drawing page is visible in Microsoft Visio Viewer. Read/write.


## Syntax

_expression_.**PageVisible**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

**Boolean**


## Remarks

The default is that the drawing page not be visible (**False**).


## Example

The following example shows how to make the drawing page visible in Visio Viewer.

```vb
vsoViewer.PageVisible = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]