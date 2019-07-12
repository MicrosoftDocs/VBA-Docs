---
title: Presentation.FarEastLineBreakLevel property (PowerPoint)
keywords: vbapp10.chm583043
f1_keywords:
- vbapp10.chm583043
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.FarEastLineBreakLevel
ms.assetid: fc8354a6-cbd4-d0b4-0b39-a3150afab714
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.FarEastLineBreakLevel property (PowerPoint)

Returns or sets the line break based upon Asian character level. Read/write.


## Syntax

_expression_. `FarEastLineBreakLevel`

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

PpFarEastLineBreakLevel


## Remarks

The value of the  **FarEastLineBreakLevel** property can be one of these **PpFarEastLineBreakLevel** constants.


||
|:-----|
|**ppFarEastLineBreakLevelCustom**|
|**ppFarEastLineBreakLevelNormal**|
|**ppFarEastLineBreakLevelStrict**|

## Example

This example sets line break control to use level one kinsoku characters.


```vb
ActivePresentation.FarEastLineBreakLevel = ppFarEastLineBreakLevelNormal
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]