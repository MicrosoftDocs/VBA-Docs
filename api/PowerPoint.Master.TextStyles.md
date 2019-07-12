---
title: Master.TextStyles property (PowerPoint)
keywords: vbapp10.chm533011
f1_keywords:
- vbapp10.chm533011
ms.prod: powerpoint
api_name:
- PowerPoint.Master.TextStyles
ms.assetid: 713b6f60-5c20-6ddf-9660-4f5f2d27546d
ms.date: 06/08/2017
localization_priority: Normal
---


# Master.TextStyles property (PowerPoint)

Returns a  **[TextStyles](PowerPoint.TextStyles.md)** collection that represents three text styles — title text, body text, and default text — for the specified slide master. Read-only.


## Syntax

_expression_. `TextStyles`

_expression_ A variable that represents a [Master](PowerPoint.Master.md) object.


## Return value

TextStyles


## Remarks

For information about returning a single member of a collection, see [Returning an object from a collection](../powerpoint/How-to/return-objects-from-collections.md).


## Example

This example sets the font name and font size for level-one body text on slides in the active presentation.


```vb
With ActivePresentation.SlideMaster_

        .TextStyles(ppBodyStyle).Levels(1)

    With .Font

        .Name = "arial"

        .Size = 36

    End With

End With
```


## See also


[Master Object](PowerPoint.Master.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]