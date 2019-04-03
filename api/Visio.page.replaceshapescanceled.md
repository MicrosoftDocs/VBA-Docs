---
title: Page.ReplaceShapesCanceled event (Visio)
ms.prod: visio
ms.assetid: 867b1fc1-96bd-cbeb-fd61-b02a96e039ca
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.ReplaceShapesCanceled event (Visio)

Occurs after an event handler has returned  **True** (cancel) to a **QueryCancelReplaceShapes** event.


## Syntax

_expression_. `ReplaceShapesCanceled`_(replaceShapes)_

 _expression_ A variable that represents a [Page](./Visio.Page.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|||||
| _replaceShapes_|Required|REPLACESHAPESEVENT|An object whose properties return information about the shape-replacement operation.|

## See also


[Page Object](Visio.Page.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]