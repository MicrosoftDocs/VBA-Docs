---
title: Viewer.ToolbarCustomizable property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.ToolbarCustomizable
ms.assetid: d49d690c-7c6d-0fab-4295-9540708eaf5c
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.ToolbarCustomizable property (Visio Viewer)

Gets or sets a value that indicates whether it is possible to customize the toolbar in Microsoft Visio Viewer. Read/write.


## Syntax

_expression_.**ToolbarCustomizable**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

**Boolean**


## Remarks

The default is for the toolbar to be customizable (**True**). When the toolbar is customizable, right-clicking the toolbar and choosing **Customize** opens the **Customize Toolbar** dialog box.


## Example

The following code makes the toolbar non-customizable in Visio Viewer.

```vb
vsoViewer.ToolbarCustomizable = False
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]