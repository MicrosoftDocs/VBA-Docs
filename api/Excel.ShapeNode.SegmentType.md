---
title: ShapeNode.SegmentType property (Excel)
keywords: vbaxl10.chm111005
f1_keywords:
- vbaxl10.chm111005
ms.prod: excel
api_name:
- Excel.ShapeNode.SegmentType
ms.assetid: 716e8171-1fd6-941e-209f-e48f5468940f
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeNode.SegmentType property (Excel)

Returns a value that indicates whether the segment associated with the specified node is straight or curved. If the specified node is a control point for a curved segment, this property returns **msoSegmentCurve**. Read-only **[MsoSegmentType](office.msosegmenttype.md)**.


## Syntax

_expression_.**SegmentType**

_expression_ A variable that represents a **[ShapeNode](Excel.ShapeNode.md)** object.


## Remarks

Use the **[SetSegmentType](Excel.ShapeNodes.SetSegmentType.md)** method to set the value of this property.


## Example

This example changes all straight segments to curved segments in shape three on _myDocument_. Shape three must be a freeform drawing.

```vb
Set myDocument = Worksheets(1) 
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


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]