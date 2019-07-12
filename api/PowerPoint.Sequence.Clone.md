---
title: Sequence.Clone method (PowerPoint)
keywords: vbapp10.chm651005
f1_keywords:
- vbapp10.chm651005
ms.prod: powerpoint
api_name:
- PowerPoint.Sequence.Clone
ms.assetid: 71dde88b-8d65-b08c-ca7b-886959fa870d
ms.date: 06/08/2017
localization_priority: Normal
---


# Sequence.Clone method (PowerPoint)

Creates a copy of an  **[Effect](PowerPoint.Effect.md)** object, and adds it to the **[Sequences](PowerPoint.Sequences.md)** collection at the specified index position.


## Syntax

_expression_. `Clone`( `_Effect_`, `_Index_` )

_expression_ A variable that represents a [Sequence](PowerPoint.Sequence.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Effect_|Required|**Effect**|**Effect** object. The animation effect to be cloned.|
| _Index_|Optional|**Long**|The position at which the cloned animation effect will be added to the  **Sequences** collection. The default value is -1 (added to the end).|

## Return value

Effect


## Example

This example copies an animation effect. This example assumes an animation effect named "effDiamond" exists.


```vb
Sub CloneEffect()
    ActivePresentation.Slides(1).TimeLine.MainSequence _
        .Clone Effect:=effDiamond, Index:=-1
End Sub
```


## See also


[Sequence Object](PowerPoint.Sequence.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]