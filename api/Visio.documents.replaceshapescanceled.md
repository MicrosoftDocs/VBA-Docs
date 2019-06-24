---
title: Documents.ReplaceShapesCanceled event (Visio)
ms.prod: visio
ms.assetid: 94a20fe7-da09-4e3c-d048-05ba0b8f1070
ms.date: 06/08/2017
localization_priority: Normal
---


# Documents.ReplaceShapesCanceled event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelReplaceShapes** event.


## Syntax

_expression_.**ReplaceShapesCanceled** (_replaceShapes_)

_expression_ A variable that represents a **[Documents](Visio.Documents.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _replaceShapes_|Required|**[REPLACESHAPESEVENT]**|An object whose properties return information about the shape-replacement operation.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]