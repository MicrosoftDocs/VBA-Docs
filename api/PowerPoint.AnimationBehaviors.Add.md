---
title: AnimationBehaviors.Add method (PowerPoint)
keywords: vbapp10.chm656004
f1_keywords:
- vbapp10.chm656004
ms.prod: powerpoint
api_name:
- PowerPoint.AnimationBehaviors.Add
ms.assetid: 427e7faa-1fc7-a145-98bc-1954054c2aec
ms.date: 06/08/2017
localization_priority: Normal
---


# AnimationBehaviors.Add method (PowerPoint)

Returns an  **[AnimationBehavior](PowerPoint.AnimationBehavior.md)** object that represents a new animation behavior.


## Syntax

_expression_.**Add** (_Type_, _Index_)

_expression_ A variable that represents an [AnimationBehaviors](PowerPoint.AnimationBehaviors.md) object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**MsoAnimType**|The type of the animation behavior.|
| _Index_|Optional|**Long**|The position of the animation behavior in relation to other animation behaviors. The default value is -1, which means that if you omit the  _Index_ parameter, the new animation behavior is added at the end of the existing animation behaviors.|

## Return value

AnimationBehavior


## Remarks

The  _Type_ parameter value can be one of these **MsoAnimType** constants.


||
|:-----|
|**msoAnimTypeColor**|
|**msoAnimTypeMixed**|
|**msoAnimTypeMotion**|
|**msoAnimTypeNone**|
|**msoAnimTypeProperty**|
|**msoAnimTypeRotation**|
|**msoAnimTypeScale**|

## See also


[AnimationBehaviors Object](PowerPoint.AnimationBehaviors.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]