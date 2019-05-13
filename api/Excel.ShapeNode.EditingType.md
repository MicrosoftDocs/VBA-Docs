---
title: ShapeNode.EditingType property (Excel)
keywords: vbaxl10.chm111003
f1_keywords:
- vbaxl10.chm111003
ms.prod: excel
api_name:
- Excel.ShapeNode.EditingType
ms.assetid: 78a17ed7-7e30-d5f3-4af8-636d65079218
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeNode.EditingType property (Excel)

If the specified node is a vertex, this property returns a value that indicates how changes made to the node affect the two segments connected to the node. Read-only **[MsoEditingType](Office.MsoEditingType.md)**.


## Syntax

_expression_.**EditingType**

_expression_ A variable that represents a **[ShapeNode](Excel.ShapeNode.md)** object.


## Remarks

This property is read-only. Use the **[SetEditingType](Excel.ShapeNodes.SetEditingType.md)** method to set the value of this property.


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