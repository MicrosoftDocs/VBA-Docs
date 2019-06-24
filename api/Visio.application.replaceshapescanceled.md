---
title: Application.ReplaceShapesCanceled event (Visio)
ms.prod: visio
ms.assetid: e8eecd64-e4bd-d2c4-b942-c5ff607a4121
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ReplaceShapesCanceled event (Visio)

Occurs after an event handler has returned **True** (cancel) to a **QueryCancelReplaceShapes** event.


## Syntax

_expression_.**ReplaceShapesCanceled** (_replaceShapes_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _replaceShapes_|Required|**[REPLACESHAPESEVENT]**|An object whose properties return information about the shape-replacement operation.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]