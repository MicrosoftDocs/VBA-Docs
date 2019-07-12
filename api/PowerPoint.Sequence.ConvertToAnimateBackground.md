---
title: Sequence.ConvertToAnimateBackground method (PowerPoint)
keywords: vbapp10.chm651010
f1_keywords:
- vbapp10.chm651010
ms.prod: powerpoint
api_name:
- PowerPoint.Sequence.ConvertToAnimateBackground
ms.assetid: 75fd5a43-f8cf-5ba9-de92-3031eb938eb7
ms.date: 06/08/2017
localization_priority: Normal
---


# Sequence.ConvertToAnimateBackground method (PowerPoint)

Determines whether the background will be animated separately from, or in addition to, its accompanying text. Returns an  **[Effect](PowerPoint.Effect.md)** object representing the newly-modified animation effect.


## Syntax

_expression_. `ConvertToAnimateBackground`( `_Effect_`, `_AnimateBackground_` )

_expression_ A variable that represents a [Sequence](PowerPoint.Sequence.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Effect_|Required|**Effect**|The animation effect to be applied to the background.|
| _AnimateBackground_|Required|**[MsoTriState](Office.MsoTriState.md)**|Determines whether the text will be animated separately from the background.|

## Return value

Effect


## Example

This example creates a text effect for the first shape on the first slide in the active presentation, and animates the text in the shape separately from the background. This example assumes there is a shape on the first slide, and that the shape has text inside it.


```vb
Sub AnimateText() 
 
    Dim timeMain As TimeLine 
    Dim shpActive As Shape 
 
    Set shpActive = ActivePresentation.Slides(1).Shapes(1) 
    Set timeMain = ActivePresentation.Slides(1).TimeLine 
 
    ' Add a blast effect to the text, and animate the text separately 
    ' from the background. 
    timeMain.MainSequence.ConvertToAnimateBackground _ 
        Effect:=timeMain.MainSequence.AddEffect(Shape:=shpActive, _ 
            effectid:=msoAnimEffectBlast), _ 
        AnimateBackGround:=msoFalse 
 
End Sub
```


## See also


[Sequence Object](PowerPoint.Sequence.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]