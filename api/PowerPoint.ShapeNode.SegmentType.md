---
title: ShapeNode.SegmentType property (PowerPoint)
keywords: vbapp10.chm561004
f1_keywords:
- vbapp10.chm561004
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeNode.SegmentType
ms.assetid: 5135d7a7-3ed7-6abd-b072-7456a59aa707
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeNode.SegmentType property (PowerPoint)

Returns a value that indicates whether the segment associated with the specified node is straight or curved. Read-only.


## Syntax

_expression_.**SegmentType**

_expression_ A variable that represents a **[ShapeNode](PowerPoint.ShapeNode.md)** object.


## Return value

MsoSegmentType


## Remarks

This property is read-only. Use the  **[SetSegmentType](PowerPoint.ShapeNodes.SetSegmentType.md)** method to set the value of this property.

The value returned by the  **SegmentType** property can be one of these **MsoSegmentType** constants. The **SegmentType** property returns **msoSegmentCurve** if the specified node is a control point for a curved segment.


||
|:-----|
|**msoSegmentCurve**|
|**msoSegmentLine**|

## Example

This example changes all straight segments to curved segments in shape three on _myDocument_. Shape three must be a freeform drawing.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(3).Nodes

    n = 1

    While n <= .Count

        If .Item(n).SegmentType = msoSegmentLine Then

            .SetSegmentType n, msoSegmentCurve

        End If

        n = n + 1

    Wend

End With
```


## See also


[ShapeNode Object](PowerPoint.ShapeNode.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]