---
title: BulletFormat.Style property (PowerPoint)
keywords: vbapp10.chm577010
f1_keywords:
- vbapp10.chm577010
ms.prod: powerpoint
api_name:
- PowerPoint.BulletFormat.Style
ms.assetid: 2cc49660-bcf7-f965-f3cb-80e6d06ba789
ms.date: 06/08/2017
localization_priority: Normal
---


# BulletFormat.Style property (PowerPoint)

Returns or sets the bullet style. Read/write. 


## Syntax

_expression_.**Style**

_expression_ A variable that represents a **[BulletFormat](PowerPoint.BulletFormat.md)** object.


## Return value

[PpNumberedBulletStyle](PowerPoint.PpNumberedBulletStyle.md)


## Remarks

Some of the  **PpNumberedBulletStyle** constants may not be available to you, depending on the language support (U.S. English, for example) that you've selected or installed.


## Example

This example sets the bullet style for the bulleted list, represented by shape one on the first slide, to a shadow color number with circular background of normal text color.


```vb
ActivePresentation.Slides(1).Shapes(1).TextFrame _
    .TextRange.ParagraphFormat.Bullet _
        .Style = ppBulletCircleNumWDBlackPlain
```


## See also


[BulletFormat Object](PowerPoint.BulletFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]