---
title: WebBrowserControl.DocumentComplete event (Access)
keywords: vbaac10.chm143141
f1_keywords:
- vbaac10.chm143141
ms.prod: access
api_name:
- Access.WebBrowserControl.DocumentComplete
ms.assetid: 8cb83f9f-b9c2-8534-8fe3-eb5c56338d6c
ms.date: 03/26/2019
localization_priority: Normal
---


# WebBrowserControl.DocumentComplete event (Access)

Occurs when a document is completely loaded and initialized.


## Syntax

_expression_.**DocumentComplete** (_pDisp_, _URL_)

_expression_ A variable that represents a **[WebBrowserControl](Access.WebBrowserControl.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pDisp_|Required|**Object**| A pointer to the **IDispatch** interface of the window or frame in which the document is loaded.|
| _URL_|Required|**Variant**|Contains the URL of the loaded document.|

## Return value

Nothing




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]