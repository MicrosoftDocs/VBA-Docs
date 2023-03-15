---
title: EdgeBrowserControl.BeforeNavigate event (Access)
keywords: vbaac10.chm143140
f1_keywords:
- vbaac10.chm143140
ms.prod: access
api_name:
- Access.EdgeBrowserControl.BeforeNavigate
ms.assetid: 616aa459-0092-470c-be41-79f99c61a020
ms.date: 03/08/2023
ms.localizationpriority: medium
---


# EdgeBrowserControl.BeforeNavigate event (Access)

Occurs before navigation occurs in the given **EdgeBrowserControl**.


## Syntax

_expression_.**BeforeNavigate** (_Cancel_, _URL_)

_expression_ A variable that represents a **[WebBrowserControl](Access.WebBrowserControl.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _URL_|Required|**String**|Contains the URL to be navigated to.|
| _Cancel_|Required|**Boolean**|Contains the cancel flag. Set to **True** to cancel the navigation operation.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]