---
title: MediaFormat.ResamplingStatus property (PowerPoint)
keywords: vbapp10.chm724015
f1_keywords:
- vbapp10.chm724015
ms.prod: powerpoint
api_name:
- PowerPoint.MediaFormat.ResamplingStatus
ms.assetid: 2a53f58e-3533-e93e-2aa1-9c6250f9c336
ms.date: 06/08/2017
localization_priority: Normal
---


# MediaFormat.ResamplingStatus property (PowerPoint)

Returns the resampling task status. Read-only.


## Syntax

_expression_. `ResamplingStatus`

 _expression_ An expression that returns a [MediaFormat](PowerPoint.MediaFormat.md) object.


## Return value

PpMediaTaskStatus


## Remarks

 **ResamplingStatus** returns one of the following **PpMediaTaskStatus** values:



|Constant|Value|Description|
|:-----|:-----|:-----|
|**ppMediaTaskStatusNone**|0|No status|
|**ppMediaTaskStatusInProgress**|1|In progress|
|**ppMediaTaskStatusQueued**|2|Queued|
|**ppMediaTaskStatusDone**|3|Done|
|**ppMediaTaskStatusFailed**|4|Failed|

## See also


[MediaFormat Object](PowerPoint.MediaFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]