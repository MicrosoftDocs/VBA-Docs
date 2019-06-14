---
title: ShapeRange.ID property (Publisher)
keywords: vbapb10.chm2293861
f1_keywords:
- vbapb10.chm2293861
ms.prod: publisher
api_name:
- Publisher.ShapeRange.ID
ms.assetid: d7ad646b-be40-2ac4-9d3e-faa37f8bf456
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeRange.ID property (Publisher)

Returns a **Long** that represents the type of a shape, range of shapes, or property, type, or value of a wizard. Read-only.


## Syntax

_expression_.**ID**

_expression_ A variable that represents a **[ShapeRange](Publisher.ShapeRange.md)** object.


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