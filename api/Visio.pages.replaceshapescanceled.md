---
title: Pages.ReplaceShapesCanceled event (Visio)
ms.prod: visio
ms.assetid: f0ce8c66-7a15-5f91-8c89-e177bc6671d2
ms.date: 06/08/2017
localization_priority: Normal
---


# Pages.ReplaceShapesCanceled event (Visio)

Occurs after an event handler has returned **True** (cancel) to a **QueryCancelReplaceShapes** event.


## Syntax

_expression_.**ReplaceShapesCanceled** (_replaceShapes_)

_expression_ A variable that represents a **[Pages](Visio.Pages.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _replaceShapes_|Required|**[REPLACESHAPESEVENT]**|An object whose properties return information about the shape-replacement operation.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]