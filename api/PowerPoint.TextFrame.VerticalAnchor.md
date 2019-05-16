---
title: TextFrame.VerticalAnchor property (PowerPoint)
keywords: vbapp10.chm558011
f1_keywords:
- vbapp10.chm558011
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame.VerticalAnchor
ms.assetid: fc38f7d8-25f7-5605-0f63-aa116fcf46d9
ms.date: 06/08/2017
localization_priority: Normal
---


# TextFrame.VerticalAnchor property (PowerPoint)

Returns or sets the vertical alignment of text in a text frame. Read/write.


## Syntax

_expression_.**VerticalAnchor**

_expression_ A variable that represents a **[TextFrame](PowerPoint.TextFrame.md)** object.


## Return value

MsoVerticalAnchor


## Remarks

The value of the  **VerticalAnchor** property can be one of these **MsoVerticalAnchor** constants.



|Constant|Description|
|:-----|:-----|
|**msoAnchorBottom**|Anchors the bottom of the text string to the current position.|
|**msoAnchorBottomBaseLine**|Anchors the bottom of the text string to the current position regardless of the resizing of text. When you resize text without baseline anchoring, the text centers itself on the previous position.|
|**msoAnchorMiddle**|Anchors the middle of the text string to the current position.|
|**msoAnchorTop**|Anchors the top of the text string to the current position|
|**msoAnchorTopBaseline**|Anchors the top of the text string to the current position regardless of the resizing of text. When you resize text without baseline anchoring, the text centers itself on the previous position.|
|**msoVerticalAnchorMixed**| Read-only. Returned when two or more text boxes within a shape range have this property set to different values.|

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