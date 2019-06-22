---
title: Window.ScrollViewTo method (Visio)
keywords: vis_sdr.chm11616525
f1_keywords:
- vis_sdr.chm11616525
ms.prod: visio
api_name:
- Visio.Window.ScrollViewTo
ms.assetid: c2930ee2-f56f-2e3e-cc9a-db73e1d99cd1
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.ScrollViewTo method (Visio)

Scrolls a window to a particular page coordinate.


## Syntax

_expression_. `ScrollViewTo`( `_x_` , `_y_` )

_expression_ A variable that represents a **[Window](Visio.Window.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _x_|Required| **Double**|The x-coordinate to which to scroll.|
| _y_|Required| **Double**|The y-coordinate to which to scroll.|

## Return value

Nothing


## Remarks

The  **ScrollViewTo** method scrolls to the _x_ and _y_ coordinates.

If the value of the  **Window** object's **Type** property is not **visDrawing**, the method raises an exception.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]