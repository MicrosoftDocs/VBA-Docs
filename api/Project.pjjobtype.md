---
title: PjJobType enumeration (Project)
ms.prod: project-server
ms.assetid: 61e64bfb-8cd8-7b76-9a5f-b7499953275f
ms.date: 06/08/2017
localization_priority: Normal
---


# PjJobType enumeration (Project)
Contains constants that specify the type of job (save, publish, or check in) that the Project Professional cache can send to the Project Server Queue System.

## Members



|Name|Value|Description|
|:-----|:-----|:-----|
|**pjCacheProjectCheckin**|1|The queue job message is to check in the project.|
|**pjCacheProjectSave**|0|The queue job message is to save the project.|
|**pjCacheProjectPublish**|2|The queue job message is to publish the project.|


## Remarks

In the **[Application.GetCacheStatusForProject](Project.application.getcachestatusforproject.md)** property, the _ProjectJobType_ parameter can be one of the **PjJobType** constants.


## See also


[GetCacheStatusForProject Property](Project.application.getcachestatusforproject.md)
[PjCacheJobState Enumeration](Project.pjcachejobstate.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]