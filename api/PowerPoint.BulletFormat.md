---
title: BulletFormat object (PowerPoint)
keywords: vbapp10.chm577000
f1_keywords:
- vbapp10.chm577000
ms.prod: powerpoint
api_name:
- PowerPoint.BulletFormat
ms.assetid: 8c70b2af-0175-9315-3a7e-e30aa0438798
ms.date: 06/08/2017
localization_priority: Normal
---


# BulletFormat object (PowerPoint)

Represents bullet formatting.


## Example

Use the [Bullet](PowerPoint.ParagraphFormat.Bullet.md)property to return the  **BulletFormat** object. The following example sets the bullet size and color for the paragraphs in shape two on slide one in the active presentation.


```vb
With ActivePresentation.Slides(1).Shapes(2)

    With .TextFrame.TextRange.ParagraphFormat.Bullet

        .Visible = True

        .RelativeSize = 1.25

        .Character = 169

        With .Font

            .Color.RGB = RGB(255, 255, 0)

            .Name = "Symbol"

        End With

    End With

End With
```


## Methods



|Name|
|:-----|
|[Picture](PowerPoint.BulletFormat.Picture.md)|

## Properties



|Name|
|:-----|
|[Application](PowerPoint.BulletFormat.Application.md)|
|[Character](PowerPoint.BulletFormat.Character.md)|
|[Font](PowerPoint.BulletFormat.Font.md)|
|[Number](PowerPoint.BulletFormat.Number.md)|
|[Parent](PowerPoint.BulletFormat.Parent.md)|
|[RelativeSize](PowerPoint.BulletFormat.RelativeSize.md)|
|[StartValue](PowerPoint.BulletFormat.StartValue.md)|
|[Style](PowerPoint.BulletFormat.Style.md)|
|[Type](PowerPoint.BulletFormat.Type.md)|
|[UseTextColor](PowerPoint.BulletFormat.UseTextColor.md)|
|[UseTextFont](PowerPoint.BulletFormat.UseTextFont.md)|

## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]