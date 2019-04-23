---
title: WebBrowserControl.BeforeNavigate2 event (Access)
keywords: vbaac10.chm143140
f1_keywords:
- vbaac10.chm143140
ms.prod: access
api_name:
- Access.WebBrowserControl.BeforeNavigate2
ms.assetid: 7f6c963b-604e-c350-e71f-899fd6258e46
ms.date: 03/26/2019
localization_priority: Normal
---


# WebBrowserControl.BeforeNavigate2 event (Access)

Occurs before navigation occurs in the given **WebBrowserControl**.


## Syntax

_expression_.**BeforeNavigate2** (_pDisp_, _URL_, _flags_, _TargetFrameName_, _PostData_, _Headers_, _Cancel_)

_expression_ A variable that represents a **[WebBrowserControl](Access.WebBrowserControl.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pDisp_|Required|**Object**|A pointer to the **IDispatch** interface for the **WebBrowserControl** object that represents the window or frame.|
| _URL_|Required|**Variant**|Contains the URL to be navigated to.|
| _flags_|Required|**Variant**|Reserved. Must be set to **Null**.|
| _TargetFrameName_|Required|**Variant**|Contains the name of the frame in which to display the resource, or **Null** if no named frame is targeted for the resource.|
| _PostData_|Required|**Variant**|Contains the data to send to the server, if the HTTP POST transaction is used.|
| _Headers_|Required|**Variant**|Contains additional HTTP headers to send to the server (HTTP URLs only). The headers can specify information, such as the action required of the server, the type of data being passed to the server, or a status code.|
| _Cancel_|Required|**Boolean**|Contains the cancel flag. Set to **True** to cancel the navigation operation.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]