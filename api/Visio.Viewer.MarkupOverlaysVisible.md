---
title: Viewer.MarkupOverlaysVisible property (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.MarkupOverlaysVisible
ms.assetid: 5e9f83b1-9c92-73b0-fa45-adf6b3ab612a
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.MarkupOverlaysVisible property (Visio Viewer)

Gets or sets a value that indicates whether markup overlays are visible in the current document open in Microsoft Visio Viewer. Read/write.


## Syntax

_expression_.**MarkupOverlaysVisible**

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Return value

**Boolean**


## Remarks

A markup overlay is a layer that shows all the shapes, ink shapes, and comments added to a drawing by a particular reviewer. The **MarkupOverlaysVisible** property setting corresponds to the status of the **Show markup overlays** box on the **Markup Settings** tab of the **Properties and Settings** dialog box in the Visio Viewer user interface. 

If markup overlays exist in the drawing, the default is for them to be visible (**True**).


## Example

The following code shows how to turn off visibility of markup overlays in Visio Viewer.

```vb
vsoViewer.MarkupOverlaysVisible = False
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]