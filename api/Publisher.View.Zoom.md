---
title: View.Zoom property (Publisher)
keywords: vbapb10.chm327684
f1_keywords:
- vbapb10.chm327684
ms.prod: publisher
api_name:
- Publisher.View.Zoom
ms.assetid: 31727291-740b-4e77-9c6b-9f19523488cb
ms.date: 06/08/2017
localization_priority: Normal
---


# View.Zoom property (Publisher)

Returns or sets a  **PbZoom** constant or a value between 10 and 400 indicating the zoom setting of the specified view. Read/write.


## Syntax

_expression_.**Zoom**

 _expression_ A variable that represents a  **View** object.


## Return value

PbZoom


## Remarks

The  **Zoom** property value can be one of the **PbZoom** constants declared in the Microsoft Publisher type library and shown in the following table.



|Constant|Description|
|:-----|:-----|
| **pbZoomFitSelection**| Resizes the page view to the size of the current selection.|
| **pbZoomPageWidth**|Resizes the page view to the width of the publication. |
| **pbZoomWholePage**| Resizes the page view to the size of a whole page.|

## Example

The following example sets the zoom for the active publication so that the entire page fits on the screen.


```vb
ActiveDocument.ActiveView.Zoom = pbZoomWholePage
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]