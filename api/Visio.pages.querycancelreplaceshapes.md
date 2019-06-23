---
title: Pages.QueryCancelReplaceShapes event (Visio)
ms.prod: visio
ms.assetid: d11ff976-0016-da6b-92fb-379baa7e8f94
ms.date: 06/08/2017
localization_priority: Normal
---


# Pages.QueryCancelReplaceShapes event (Visio)

Occurs immediately after a shape-replacement operation is requested. If any event handler returns  **True**, the operation is canceled.


## Syntax

_expression_.**QueryCancelReplaceShapes** ( _replaceShapes_, _lpboolRet_)

_expression_ A variable that represents a **[Pages](Visio.Pages.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _replaceShapes_|Required|**[REPLACESHAPESEVENT]**|An object whose properties return information about the shape-replacement operation.|
| _lpboolRet_|Required|BOOL||



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]