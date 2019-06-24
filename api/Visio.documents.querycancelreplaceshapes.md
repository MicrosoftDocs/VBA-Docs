---
title: Documents.QueryCancelReplaceShapes event (Visio)
ms.prod: visio
ms.assetid: d613730e-04c8-d17f-0ad1-19e976aa107d
ms.date: 06/08/2017
localization_priority: Normal
---


# Documents.QueryCancelReplaceShapes event (Visio)

Occurs immediately after a shape-replacement operation is requested. If any event handler returns  **True**, the operation is canceled.


## Syntax

_expression_.**QueryCancelReplaceShapes** (_replaceShapes_, _lpboolRet_)

_expression_ A variable that represents a **[Documents](Visio.Documents.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _replaceShapes_|Required|**[REPLACESHAPESEVENT]**|An object whose properties return information about the shape-replacement operation.|
| _lpboolRet_|Required|BOOL||



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]