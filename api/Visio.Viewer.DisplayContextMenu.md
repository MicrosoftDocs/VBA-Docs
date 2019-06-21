---
title: Viewer.DisplayContextMenu method (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.DisplayContextMenu
ms.assetid: 0aa19901-7bb8-6abe-cbff-4217381af336
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.DisplayContextMenu method (Visio Viewer)

Displays the shortcut menu for Microsoft Visio Viewer at the specified screen coordinates, in pixels.


## Syntax

_expression_.**DisplayContextMenu** (_ScreenX_, _ScreenY_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ScreenX_|Required| **Long**|The x-coordinate, in pixels, of the point where the menu should appear, relative to the origin of the frame of the screen.|
|_ScreenY_|Required| **Long**|The y-coordinate, in pixels, of the point where the menu should appear, relative to the origin of the frame of the screen.|

## Return value

Nothing


## Remarks

Use the _ScreenX_ and _ScreenY_ parameters to specify the coordinates of the point where you want the shortcut menu to appear, relative to the origin of the frame of the screen. The origin of the screen frame is in the upper-left corner.


## Example

The following code specifies that the shortcut menu appear at screen coordinates (300, 300).

```vb
vsoViewer.DisplayContextMenu(300,300)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]