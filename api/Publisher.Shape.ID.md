---
title: Shape.ID property (Publisher)
keywords: vbapb10.chm2228325
f1_keywords:
- vbapb10.chm2228325
ms.prod: publisher
api_name:
- Publisher.Shape.ID
ms.assetid: df4ccd93-e3fa-eeef-b5ea-e99aa0dde199
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.ID property (Publisher)

Returns a **Long** that represents the type of a shape, range of shapes, or property, type, or value of a wizard. Read-only.


## Syntax

_expression_.**ID**

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Example

This example displays the type for each shape on the first page of the active publication.

```vb
Sub ShapeID() 
 Dim shp As Shape 
 For Each shp In ActiveDocument.Pages(1).Shapes 
 MsgBox shp.ID 
 Next shp 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]