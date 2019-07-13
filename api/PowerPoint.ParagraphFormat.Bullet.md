---
title: ParagraphFormat.Bullet property (PowerPoint)
keywords: vbapp10.chm576004
f1_keywords:
- vbapp10.chm576004
ms.prod: powerpoint
api_name:
- PowerPoint.ParagraphFormat.Bullet
ms.assetid: 2b997a78-7791-6f08-00af-7143f94457c1
ms.date: 06/08/2017
localization_priority: Normal
---


# ParagraphFormat.Bullet property (PowerPoint)

Returns a  **[BulletFormat](PowerPoint.BulletFormat.md)** object that represents bullet formatting for the specified paragraph format. Read-only.


## Syntax

_expression_. `Bullet`

_expression_ A variable that represents a [ParagraphFormat](PowerPoint.ParagraphFormat.md) object.


## Return value

BulletFormat


## Example

This example sets the bullet size and bullet color for the paragraphs in shape two on slide one in the active presentation.


```vb
With Application.ActivePresentation.Slides(1).Shapes(2).TextFrame

    With .TextRange.ParagraphFormat.Bullet

        .Visible = True

        .RelativeSize = 1.25

        .Font.Color = RGB(255, 0, 255)

    End With

End With
```


## See also


[ParagraphFormat Object](PowerPoint.ParagraphFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
