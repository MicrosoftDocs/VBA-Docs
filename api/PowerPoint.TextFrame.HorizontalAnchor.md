---
title: TextFrame.HorizontalAnchor property (PowerPoint)
keywords: vbapp10.chm558010
f1_keywords:
- vbapp10.chm558010
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame.HorizontalAnchor
ms.assetid: 9f694882-ce8d-d412-d60e-5217e92a81a7
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame.HorizontalAnchor property (PowerPoint)

Returns or sets the horizontal alignment of text in a text frame. Read/write.


## Syntax

_expression_.**HorizontalAnchor**

_expression_ A variable that represents a **[TextFrame](PowerPoint.TextFrame.md)** object.


## Return value

MsoHorizontalAnchor


## Remarks

The value returned by the  **HorizontalAnchor** property can be one of these **MsoHorizontalAnchor** constants.


||
|:-----|
|**msoAnchorNone**|
|**msoHorizontalAnchorMixed**|
|**msoAnchorCenter**|

## Example

This example sets the alignment of the text in shape one on _myDocument_ to top centered.


```vb
Set myDocument = ActivePresentation.SlideMaster

With myDocument.Shapes(1)

    .TextFrame.HorizontalAnchor = msoAnchorCenter

    .TextFrame.VerticalAnchor = msoAnchorTop

End With
```


## See also


[TextFrame Object](PowerPoint.TextFrame.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]