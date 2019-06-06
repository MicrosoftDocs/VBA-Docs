---
title: Document.ShapesAdded event (Publisher)
keywords: vbapb10.chm285212675
f1_keywords:
- vbapb10.chm285212675
ms.prod: publisher
api_name:
- Publisher.Document.ShapesAdded
ms.assetid: f6573f7c-56fa-1efa-9dba-39cde3859cc0
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.ShapesAdded event (Publisher)

Occurs when one or more new shapes are added to a publication. This event occurs whether shapes are added manually or programmatically.


## Syntax

_expression_.**ShapesAdded**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Example

This example displays a message whenever a new shape is added to the active publication. For this example to work, you must place this code into the **ThisDocument** module.

```vb
Private Sub PubDoc_ShapesAdded() 
 MsgBox "You just added a new shape." 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]