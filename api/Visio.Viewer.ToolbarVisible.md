---
title: Viewer.ToolbarVisible property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.ToolbarVisible
ms.assetid: 55e6b5fc-bda6-fff4-9049-b4aa398a4744
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.ToolbarVisible property (Visio Viewer)

Gets or sets a value that indicates whether the toolbar is visible in Microsoft Visio Viewer. Read/write.


## Syntax

_expression_.**ToolbarVisible**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

**Boolean**


## Remarks

The default is for the toolbar to be visible (**True**).


## Example

The following code hides the toolbar in Visio Viewer.

```vb
vsoViewer.ToolbarVisible = False
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]