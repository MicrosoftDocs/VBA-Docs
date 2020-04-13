---
title: Sequence.AddTriggerEffect method (PowerPoint)
keywords: vbapp10.chm651013
f1_keywords:
- vbapp10.chm651013
ms.prod: powerpoint
api_name:
- PowerPoint.Sequence.AddTriggerEffect
ms.assetid: 65acf575-5b64-e95c-827d-dada8e915666
ms.date: 06/08/2017
localization_priority: Normal
---


# Sequence.AddTriggerEffect method (PowerPoint)

Adds a trigger effect to the animation in a **Sequence** object.


## Syntax

_expression_. `AddTriggerEffect`( `_pShape_`, `_effectId_`, `_trigger_`, `_pTriggerShape_`, `_bookmark_`, `_Level_` )

_expression_ A variable that represents a [Sequence](PowerPoint.Sequence.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pShape_|Required|**Shape**|The **Shape** object with animation.|
| _effectId_|Required|**MsoAnimEffect**|The type of animation.|
| _trigger_|Required|**MsoAnimTriggerType**|The type of trigger effect to add.|
| _pTriggerShape_|Required|**Shape**|The **Shape** object that represents the trigger.|
| _bookmark_|Optional|**String**|The bookmark.|
| _Level_|Optional|**MsoAnimateByLevel**|The level of animation.|

## Return value

Effect


## See also


[Sequence Object](PowerPoint.Sequence.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]