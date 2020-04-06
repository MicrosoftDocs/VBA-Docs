---
title: DocumentWindow.PointsToScreenPixelsY method (PowerPoint)
keywords: vbapp10.chm511028
f1_keywords:
- vbapp10.chm511028
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow.PointsToScreenPixelsY
ms.assetid: 0a5a96c6-3e91-31c6-ee60-ca1f8481daf0
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentWindow.PointsToScreenPixelsY method (PowerPoint)

Converts a vertical measurement from points to pixels. Used to return a vertical screen location for a text frame or shape. Returns the converted measurement as a  **Single**.


## Syntax

_expression_.**PointsToScreenPixelsY** (_Points_)

_expression_ A variable that represents a [DocumentWindow](PowerPoint.DocumentWindow.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Points_|Required|**Single**|The vertical measurement (in points) to be converted to pixels.|

## Return value

Single


## Example

This example converts the width and height of the selected text frame bounding box from points to pixels, and returns the values to  `myXparm` and `myYparm`.


```vb
With ActiveWindow

    myXparm = .PointsToScreenPixelsX _
        (.Selection.TextRange.BoundWidth)

    myYparm = .PointsToScreenPixelsY _
        (.Selection.TextRange.BoundHeight)

End With
```


## See also



[DocumentWindow Object](PowerPoint.DocumentWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]