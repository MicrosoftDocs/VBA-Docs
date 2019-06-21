---
title: Viewer.PageTabsVisible property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.PageTabsVisible
ms.assetid: 7ca92d5f-2d34-93f6-a5ca-b331125a847f
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.PageTabsVisible property (Visio Viewer)

Gets or sets a value that indicates whether page tabs are visible in Microsoft Visio Viewer. Read/write.


## Syntax

_expression_.**PageTabsVisible**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

**Boolean**


## Remarks

The default is for the page tabs not to be visible (**False**).


## Example

The following code makes the page tabs visible in Visio Viewer.

```vb
vsoViewer.PageTabsVisible = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]