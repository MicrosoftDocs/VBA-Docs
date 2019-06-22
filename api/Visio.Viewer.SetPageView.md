---
title: Viewer.SetPageView method (Visio Viewer)
ms.prod: visio
api_name:
- Visio.Viewer.SetPageView
ms.assetid: 669c8d29-9793-08a3-05ee-54aab77881bb
ms.date: 06/21/2019
localization_priority: Normal
---


# Viewer.SetPageView method (Visio Viewer)

Sets the position and zoom factor (size) of the drawing page in Microsoft Visio Viewer.


## Syntax

_expression_.**SetPageView** (_PageXAtViewCenter_, _PageYAtViewCenter_, _ZoomFactor_)

_expression_ An expression that returns a **[Viewer](Visio.Viewer.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_PageXAtViewCenter_|Required| **Double**|The x-coordinate, in drawing-page units, of the center of the page, measured from the lower-left corner of the page.|
|_PageYAtViewCenter_|Required| **Double**|The y-coordinate, in drawing-page units, of the center of the page, measured from the lower-left corner of the page.|
|_ZoomFactor_|Required| **Double**|The factor by which to multiply the zoom (page size).|

## Return value

Nothing


## Remarks

The page view consists of the center point of the page, expressed in x-y page coordinates, with the origin of the coordinate system at the lower-left corner of the page, and the zoom factor, expressed as a numerical percentage, with a range from 1% through 400%.

You can use the **[GetPageView](Visio.Viewer.GetPageView.md)** method to get the current page-view values.

The **SetPageView** method sets the coordinates of the point in the page coordinate system that is at the center of the Visio Viewer window. For example, passing 0 for both the x-coordinate and y-coordinate places the lower-left corner of the page (the origin of the page's coordinate system) in the center of the Visio Viewer window. 

If the page is 8 page-units wide by 10 page-units high, passing 4 for _PageXAtViewCenter_ and 5 for _PageYAtViewCenter_ places the center of the page at the center of the Visio Viewer window.

The _ZoomFactor_ parameter value is the factor by which to multiply both dimensions of the page. For example, passing .50 for _ZoomFactor_ makes the page both half as high and half as wide as it was previously.


## Example

The following code sets the center of the page at the center of the Visio Viewer window and halves both the height and width of the page.

```vb
vsoViewer.SetPageView 4, 5, 0.50
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]