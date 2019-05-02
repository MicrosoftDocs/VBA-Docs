---
title: BulletFormat.StartValue property (PowerPoint)
keywords: vbapp10.chm577011
f1_keywords:
- vbapp10.chm577011
ms.prod: powerpoint
api_name:
- PowerPoint.BulletFormat.StartValue
ms.assetid: d243b5b4-93f6-8486-d432-a91a39ee4f81
ms.date: 06/08/2017
localization_priority: Normal
---


# BulletFormat.StartValue property (PowerPoint)

Returns or sets the beginning value of a bulleted list when the  **[Type](PowerPoint.BulletFormat.Type.md)** property of the **BulletFormat** object is set to **ppBulletNumbered**. Read/write.


## Syntax

_expression_. `StartValue`

_expression_ A variable that represents a **[BulletFormat](PowerPoint.BulletFormat.md)** object.


## Return value

Integer


## Remarks

The value of the  **StartValue** property must be in the range of 1 to 32767.


## Example

This example sets the bullets in the text box specified by shape two on slide one to start with the number five.


```vb
With ActivePresentation.Slides(1).Shapes(2).TextFrame

    With .TextRange.ParagraphFormat.Bullet

        .Type = ppBulletNumbered

        .StartValue = 5

    End With

End With


```


## See also


[BulletFormat Object](PowerPoint.BulletFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]