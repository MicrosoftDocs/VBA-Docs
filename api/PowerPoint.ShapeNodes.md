---
title: ShapeNodes object (PowerPoint)
keywords: vbapp10.chm560000
f1_keywords:
- vbapp10.chm560000
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeNodes
ms.assetid: 493bacfe-eb8c-2064-46ec-c19e58e9b1ce
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeNodes object (PowerPoint)

A collection of all the  **[ShapeNode](PowerPoint.ShapeNode.md)** objects in the specified freeform.


## Remarks

 Each **ShapeNode** object represents either a node between segments in a freeform or a control point for a curved segment of a freeform. You can create a freeform manually or by using the [BuildFreeform](PowerPoint.Shapes.BuildFreeform.md)and [ConvertToShape](PowerPoint.FreeformBuilder.ConvertToShape.md)methods.


## Example

Use the  **Nodes** property to return the **ShapeNodes** collection. The following example deletes node four in shape three on _myDocument_. For this example to work, shape three must be a freeform with at least four nodes.


```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(3).Nodes.Delete 4
```

Use the [Insert](PowerPoint.ShapeNodes.Insert.md)method to create a new node and add it to the  **ShapeNodes** collection. The following example adds a smooth node with a curved segment after node four in shape three on _myDocument_. For this example to work, shape three must be a freeform with at least four nodes.




```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3).Nodes

    .Insert 4, msoSegmentCurve, msoEditingSmooth, 210, 100

End With
```

Use  **Nodes** (_index_), where _index_ is the node index number, to return a single **ShapeNode** object. If node one in shape three on _myDocument_ is a corner point, the following example makes it a smooth point. For this example to work, shape three must be a freeform.




```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3)

    If .Nodes(1).EditingType = msoEditingCorner Then

        .Nodes.SetEditingType 1, msoEditingSmooth

    End If

End With
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]