---
title: Application.QueryCancelReplaceShapes Event (Visio)
ms.prod: visio
ms.assetid: 50c0f2a6-f534-f3af-7e83-c865abda8bf9
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.QueryCancelReplaceShapes Event (Visio)

Occurs immediately after a shape-replacement operation is requested. If any event handler returns  **True** , the operation is canceled.


## Syntax

 _expression_. `QueryCancelReplaceShapes`( _replaceShapes_)

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|||||
| _replaceShapes_|Required|REPLACESHAPESEVENT|An object whose properties return information about the shape-replacement operation.|
|||||
| _lpboolRet_|Required|BOOL||

## See also


[Application Object](Visio.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]