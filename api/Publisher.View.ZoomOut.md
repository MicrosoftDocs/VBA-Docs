---
title: View.ZoomOut method (Publisher)
keywords: vbapb10.chm327687
f1_keywords:
- vbapb10.chm327687
ms.prod: publisher
api_name:
- Publisher.View.ZoomOut
ms.assetid: 5066a532-03a9-9b2a-b254-a1388c35bc79
ms.date: 06/15/2019
localization_priority: Normal
---


# View.ZoomOut method (Publisher)

Decreases the magnification of the specified view.


## Syntax

_expression_.**ZoomOut**

_expression_ A variable that represents a **[View](Publisher.View.md)** object.


## Example

This example decreases the magnification of the active view.

```vb
Sub Zoom() 
 ActiveView.ZoomOut 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]