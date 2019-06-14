---
title: ShapeNodes.Delete method (Publisher)
keywords: vbapb10.chm3473425
f1_keywords:
- vbapb10.chm3473425
ms.prod: publisher
api_name:
- Publisher.ShapeNodes.Delete
ms.assetid: 09f7a8ef-cefd-5a68-f0a6-e99c2f111ea6
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeNodes.Delete method (Publisher)

Deletes the specified shape node object.


## Syntax

_expression_.**Delete** (_Index_)

_expression_ A variable that represents a **[ShapeNodes](Publisher.ShapeNodes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_|Required| **Integer**/**Long**| The number of the shape node to delete.|


## Example

This example deletes the first node in the first shape in the active publication.

```vb
Sub DeleteNode() 
 ActiveDocument.Pages(1).Shapes(1).Nodes.Delete Index:=1 
End Sub
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]