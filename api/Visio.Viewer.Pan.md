---
title: Viewer.Pan method (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.Pan
ms.assetid: 5cfeabcd-37fa-ade7-2fe0-b1e19259f6cd
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.Pan method (Visio Viewer)

Moves the page by the specified coordinate values, in pixels, in Microsoft Visio Viewer. 


## Syntax

_expression_.**Pan** (_DeltaX_, _DeltaY_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_DeltaX_|Required| **Long**|The amount, in pixels, to move horizontally.|
|_DeltaY_|Required| **Long**|The amount, in pixels, to move vertically.|

## Return value

Nothing


## Remarks

The values of _DeltaX_ and _DeltaY_ can be positive or negative.


## Example

The following code moves the page 100 pixels to the right (horizontally) and 200 pixels down (vertically).

```vb
vsoViewer.Pan 100, 200
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]