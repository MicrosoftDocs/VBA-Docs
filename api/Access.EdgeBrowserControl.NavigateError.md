---
title: EdgeBrowserControl.NavigateError event (Access)
keywords: vbaac10.chm143143
f1_keywords:
- vbaac10.chm143143
ms.prod: access
api_name:
- Access.EdgeBrowserControl.NavigateError
ms.assetid: f311aed0-a85d-446f-bb1f-db88724e25eb
ms.date: 03/08/2023
ms.localizationpriority: medium
---


# EdgeBrowserControl.NavigateError event (Access)

Occurs when an error occurs during navigation.


## Syntax

_expression_.**NavigateError** (_URL_, _StatusCode_)

_expression_ A variable that represents a **[EdgeBrowserControl](Access.EdgeBrowserControl.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _URL_|Required|**String**|Contains the URL for which navigation failed.|
| _StatusCode_|Required|**Variant**|Contains an error status code, if available.|

## Status Codes
|Code|Name|
|:-----|:-----|
|0|Success|
|1|Redirected|
|2|Cannot Connect|
|3|Disconnected|
|4|Connection Aborted|
|5|Timeout|
|6|Connection Reset|
|7|Server Unreachable|
|8|Host Name Not Resolved|
|9|Certificate Common Name Is Incorrect|
|10|Operation Cancelled|
|11|Redirect Failed|
|12|Certificate Expired|
|13|Client Certificate Contains Errors|
|14|Certificate Revoked|
|15|Certificate Is Invalid|
|16|Error Http Invalid Server Response|
|17|Unexpected Error|

## Return value

Nothing






[!include[Support and feedback](~/includes/feedback-boilerplate.md)]