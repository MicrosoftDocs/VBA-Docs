---
title: Application.ReplaceShapesCanceled Event (Visio)
ms.prod: visio
ms.assetid: e8eecd64-e4bd-d2c4-b942-c5ff607a4121
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ReplaceShapesCanceled Event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelReplaceShapes** event.


## Syntax

 _expression_. `ReplaceShapesCanceled`_(replaceShapes)_

 _expression_ A variable that represents a [Application](./Visio.Application.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|||||
| _replaceShapes_|Required|REPLACESHAPESEVENT|An object whose properties return information about the shape-replacement operation.|

## See also


[Application Object](Visio.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]