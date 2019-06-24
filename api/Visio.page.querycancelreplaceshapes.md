---
title: Page.QueryCancelReplaceShapes event (Visio)
ms.prod: visio
ms.assetid: 17ead23f-825a-c608-3315-e2eed6784cd5
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.QueryCancelReplaceShapes event (Visio)

Occurs immediately after a shape-replacement operation is requested. If any event handler returns  **True**, the operation is canceled.


## Syntax

_expression_.**QueryCancelReplaceShapes** ( _replaceShapes_, _lpboolRet_)

_expression_ A variable that represents a **[Page](Visio.Page.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _replaceShapes_|Required|**[REPLACESHAPESEVENT]**|An object whose properties return information about the shape-replacement operation.|
| _lpboolRet_|Required|BOOL||



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]