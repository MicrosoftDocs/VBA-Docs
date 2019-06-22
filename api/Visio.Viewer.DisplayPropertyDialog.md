---
title: Viewer.DisplayPropertyDialog method (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.DisplayPropertyDialog
ms.assetid: 92578d7a-53a1-0597-e4b6-21444db0dad8
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.DisplayPropertyDialog method (Visio Viewer)

Displays the **Properties and Settings** dialog box at the specified screen coordinates, in pixels, in Microsoft Visio Viewer.


## Syntax

_expression_.**DisplayPropertyDialog** (_ScreenX_, _ScreenY_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ScreenX_|Optional| **Variant**|The x-coordinate, in pixels, of the point where the dialog box should appear, relative to the origin of the frame of the screen.|
|_ScreenY_|Optional| **Variant**|The y-coordinate, in pixels, of the point where the dialog box should appear, relative to the origin of the frame of the screen.|

## Return value

Nothing


## Remarks

Use the optional _ScreenX_ and _ScreenY_ parameters to specify the coordinates of the point where you want the dialog box to appear, relative to the origin of the frame of the screen. The origin of the screen frame is in the upper-left corner. If you do not specify coordinates, the dialog box appears in its default position, at the lower-right corner of the Visio Viewer control.


## Example

The following code displays the **Properties and Settings** dialog box at screen coordinates (300, 300).

```vb

Dim lngScreenPosX As Long 

Dim lngScreenPosY As Long

lngScreenPosX = 300

lngScreenPosY = 300 

vsoViewer.DisplayPropertyDialog lngScreenPosX, lngScreenPosY

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]