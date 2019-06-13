---
title: Shape.LinkFormat property (Publisher)
keywords: vbapb10.chm2228326
f1_keywords:
- vbapb10.chm2228326
ms.prod: publisher
api_name:
- Publisher.Shape.LinkFormat
ms.assetid: 801c3a87-7cc6-8c7b-094a-55e8d8d7a004
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.LinkFormat property (Publisher)

Returns a **[LinkFormat](Publisher.LinkFormat.md)** object that contains the properties that are unique to linked OLE objects. Read-only.


## Syntax

_expression_.**LinkFormat**

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Example

This example updates the links between any OLE objects on page one in the active publication and their source files.

```vb
Dim sh As Shape 
 
For Each sh In ActiveDocument.Pages(1).Shapes 
 If sh.Type = pbLinkedOLEObject Then 
 With sh.LinkFormat 
 .Update 
 End With 
 End If 
Next
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]