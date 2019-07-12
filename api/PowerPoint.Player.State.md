---
title: Player.State property (PowerPoint)
keywords: vbapp10.chm726009
f1_keywords:
- vbapp10.chm726009
ms.prod: powerpoint
api_name:
- PowerPoint.Player.State
ms.assetid: 927216b3-54b7-b00c-9812-ac274bfa5348
ms.date: 06/08/2017
localization_priority: Normal
---


# Player.State property (PowerPoint)

Returns the current state of the player. Read-only.


## Syntax

_expression_. `State`

_expression_ A variable that represents a [Player](PowerPoint.Player.md) object.


## Remarks

 **State** can return the following **PpPlayerState** values.



|Constant|Value|Description|
|:-----|:-----|:-----|
|**ppPlaying**|0|Playing|
|**ppPaused**|1|Paused|
|**ppStopped**|2|Stopped|
|**ppNotReady**|3|Not ready|

## See also


[Player Object](PowerPoint.Player.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]