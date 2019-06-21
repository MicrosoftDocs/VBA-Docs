---
title: Viewer.PropertyDialogEnabled property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.PropertyDialogEnabled
ms.assetid: 66055cb8-535d-16e5-386d-1e7a44faa669
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.PropertyDialogEnabled property (Visio Viewer)

Gets or sets a value that indicates whether the **Properties and Settings** dialog box is available in the user interface for Microsoft Visio Viewer. Read/write.


## Syntax

_expression_.**PropertyDialogEnabled**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

**Boolean**


## Remarks

The default is for the **Properties and Settings** dialog box to be available (**True**).

When the **PropertyDialogEnabled** property is set to **False**, choosing **Properties and Settings** on the toolbar or on the shortcut (right-click) menu has no effect.


## Example

The following code gets a value that indicates whether the **Properties and Settings** dialog box is available in Visio Viewer.

```vb
Debug.Print vsoViewer.PropertyDialogEnabled
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]