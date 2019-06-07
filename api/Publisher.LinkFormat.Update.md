---
title: LinkFormat.Update method (Publisher)
keywords: vbapb10.chm4390916
f1_keywords:
- vbapb10.chm4390916
ms.prod: publisher
api_name:
- Publisher.LinkFormat.Update
ms.assetid: a167a463-56bd-2c4e-ded5-70ea38b2ed2f
ms.date: 06/08/2019
localization_priority: Normal
---


# LinkFormat.Update method (Publisher)

Updates the specified linked OLE object.


## Syntax

_expression_.**Update**

_expression_ A variable that represents a **[LinkFormat](Publisher.LinkFormat.md)** object.


## Example

This example updates all linked OLE objects in the active publication.

```vb
Dim pageLoop As Page 
Dim shpLoop As Shape 
 
For Each pageLoop In ActiveDocument.Pages 
 For Each shpLoop In pageLoop.Shapes 
 
 With shpLoop 
 If .Type = pbLinkedOLEObject Then 
 .LinkFormat.Update 
 End If 
 End With 
 
 Next shpLoop 
Next pageLoop
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]