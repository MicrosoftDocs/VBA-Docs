---
title: Viewer.Pan Method (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Pan
ms.assetid: 5cfeabcd-37fa-ade7-2fe0-b1e19259f6cd
ms.date: 06/08/2017
localization_priority: Normal
---


# Viewer.Pan Method (Visio Viewer)

Moves the page by the specified coordinate values, in pixels, in Microsoft Visio Viewer. 


## Syntax

_expression_.**Pan**(**_DeltaX_**,  **_DeltaY_**)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|DeltaX|Required| **Long**|The amount, in pixels, to move horizontally.|
|DeltaY|Required| **Long**|The amount, in pixels, to move vertically.|

## Return value

Nothing


## Remarks

The values of DeltaX and DeltaY can be positive or negative.


## Example

The following code moves the page 100 pixels to the right (horizontally) and 200 pixels down (vertically).


```vb
vsoViewer.Pan 100, 200
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]