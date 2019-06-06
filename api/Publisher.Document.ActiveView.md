---
title: Document.ActiveView property (Publisher)
keywords: vbapb10.chm196707
f1_keywords:
- vbapb10.chm196707
ms.prod: publisher
api_name:
- Publisher.Document.ActiveView
ms.assetid: 1448c8c6-30e5-2e2a-f124-ebf544d8f297
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.ActiveView property (Publisher)

Returns a **[View](Publisher.View.md)** object representing the view attributes for the specified document. Read-only.


## Syntax

_expression_.**ActiveView**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Return value

View


## Example

The following example sets the active publication zoom to fill the screen.

```vb
Sub SetActiveZoom() 
 Dim viewTemp As View 
 
 ActiveDocument.Pages(1).Shapes.AddShape 1, 10, 10, 50, 50 
 Set viewTemp = ActiveDocument.ActiveView 
 ActiveDocument.Pages(1).Shapes(1).Select 
 viewTemp.Zoom = pbZoomFitSelection 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]