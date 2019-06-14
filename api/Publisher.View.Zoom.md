---
title: View.Zoom property (Publisher)
keywords: vbapb10.chm327684
f1_keywords:
- vbapb10.chm327684
ms.prod: publisher
api_name:
- Publisher.View.Zoom
ms.assetid: 31727291-740b-4e77-9c6b-9f19523488cb
ms.date: 06/15/2019
localization_priority: Normal
---


# View.Zoom property (Publisher)

Returns or sets a **[PbZoom](publisher.pbzoom.md)** constant or a value between 10 and 400 indicating the zoom setting of the specified view. Read/write.


## Syntax

_expression_.**Zoom**

_expression_ A variable that represents a **[View](Publisher.View.md)** object.


## Return value

PbZoom


## Remarks

The **Zoom** property value can be one of the **PbZoom** constants declared in the Microsoft Publisher type library.


## Example

The following example sets the zoom for the active publication so that the entire page fits on the screen.

```vb
ActiveDocument.ActiveView.Zoom = pbZoomWholePage
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]