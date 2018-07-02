---
title: TextFrame2.VerticalAnchor Property (Office)
ms.prod: office
api_name:
- Office.TextFrame2.VerticalAnchor
ms.assetid: b61506d6-05d9-84af-dd1c-3e7ebd2ea92b
ms.date: 06/08/2017
---


# TextFrame2.VerticalAnchor Property (Office)

Returns or sets the vertical alignment of text in a text frame. Read/write


## Syntax

 _expression_. `VerticalAnchor`

 _expression_ An expression that returns a [TextFrame2](./Office.TextFrame2.md) object.


## Remarks

The value of the VerticalAnchor property can be one of these MsoVerticalAnchor constants.


## Example

The following example shows how to set the alignment for shape one on slide one to top center.


```vb
With ActivePresentation.Slides(1).Shapes(1) 
 .TextFrame2.HorizontalAnchor = msoAnchorCenter 
 .TextFrame2.VerticalAnchor = msoAnchorTop 
End With
```


## See also


[TextFrame2 Object](Office.TextFrame2.md)



[TextFrame2 Object Members](./overview/textframe2-members-office.md)

