---
title: TextFrame2.HorizontalAnchor property (Office)
ms.prod: office
api_name:
- Office.TextFrame2.HorizontalAnchor
ms.assetid: 27419e1a-63e6-a08b-2d45-0cd21ada8889
ms.date: 01/25/2019
localization_priority: Normal
---


# TextFrame2.HorizontalAnchor property (Office)

Returns or sets the horizontal alignment of text in a text frame. Read/write.


## Syntax

_expression_.**HorizontalAnchor**

_expression_ An expression that returns a **[TextFrame2](Office.TextFrame2.md)** object.


## Remarks

The value of the **HorizontalAnchor** property can be one of these **[MsoHorizontalAnchor](office.msohorizontalanchor.md)** constants:

- **msoAnchorNone**
- **msoHorizontalAnchorMixed**
- **msoAnchorCenter**

## Example

The following code shows how to set the alignment for shape one on slide one to top center.

```vb
With ActivePresentation.Slides(1).Shapes(1) 
 .TextFrame2.HorizontalAnchor = msoAnchorCenter 
 .TextFrame2.VerticalAnchor = msoAnchorTop 
End With
```


## See also

- [TextFrame2 object members](overview/Library-Reference/textframe2-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]