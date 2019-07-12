---
title: Presentation.NoLineBreakBefore property (PowerPoint)
keywords: vbapp10.chm583044
f1_keywords:
- vbapp10.chm583044
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.NoLineBreakBefore
ms.assetid: d7f7f559-cf20-ef3f-60aa-122dc28da203
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.NoLineBreakBefore property (PowerPoint)

Returns or sets the characters that cannot begin a line. Read/write.


## Syntax

_expression_. `NoLineBreakBefore`

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

String


## Example

This example sets "!", ")", and "]" as characters that cannot begin a line.


```vb
With ActivePresentation

    .FarEastLineBreakLevel = ppFarEastLineBreakLevelCustom

    .NoLineBreakBefore =  "!)]"

End With
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]