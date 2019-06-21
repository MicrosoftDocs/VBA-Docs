---
title: Viewer.ScrollbarsVisible property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.ScrollbarsVisible
ms.assetid: cd8f5b2d-f604-8bac-2e82-338cfa7d7174
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.ScrollbarsVisible property (Visio Viewer)

Gets or sets a value that indicates whether the scroll bars are visible in Microsoft Visio Viewer. Read/write.


## Syntax

_expression_.**ScrollbarsVisible**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

**Boolean**


## Remarks

The default is for the scroll bars to be visible (**True**).


## Example

The following code turns off display of the scroll bars in the drawing that is open in Visio Viewer.

```vb
vsoViewer.ScrollbarsVisible = False
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]