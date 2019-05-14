---
title: ShapeNodes.SetEditingType method (Excel)
keywords: vbaxl10.chm112009
f1_keywords:
- vbaxl10.chm112009
ms.prod: excel
api_name:
- Excel.ShapeNodes.SetEditingType
ms.assetid: 5bf464d6-b9d3-f62b-a625-0d153d7f265e
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeNodes.SetEditingType method (Excel)

Sets the editing type of the node specified by _Index_. If the node is a control point for a curved segment, this method sets the editing type of the node adjacent to it that joins two segments. Note that, depending on the editing type, this method may affect the position of adjacent nodes.


## Syntax

_expression_.**SetEditingType** (_Index_, _EditingType_)

_expression_ A variable that represents a **[ShapeNodes](Excel.ShapeNodes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The node whose editing type is to be set.|
| _EditingType_|Required| **[MsoEditingType](Office.MsoEditingType.md)**|The editing property of the vertex.|

## Example

This example changes all corner nodes to smooth nodes in shape three on _myDocument_. Shape three must be a freeform drawing.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3).Nodes 
    For n = 1 to .Count 
        If .Item(n).EditingType = msoEditingCorner Then 
            .SetEditingType n, msoEditingSmooth 
        End If 
    Next 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]