---
title: Application.ProtectedViewWindowBeforeClose event (PowerPoint)
keywords: vbapp10.chm621028
f1_keywords:
- vbapp10.chm621028
ms.prod: powerpoint
api_name:
- PowerPoint.Application.ProtectedViewWindowBeforeClose
ms.assetid: e10ffe16-aad8-1e2d-fd75-82243a56ef05
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProtectedViewWindowBeforeClose event (PowerPoint)

Occurs immediately before a Protected View window or a document in a Protected View window closes.


## Syntax

_expression_. `ProtectedViewWindowBeforeClose`( `_ProtViewWindow_`, `_ProtectedViewCloseReason_`, `_Cancel_` )

_expression_ A variable that represents an **[Application](PowerPoint.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ProtViewWindow_|Required|**ProtectedViewWindow**|The Protected View window that is closed.|
| _ProtectedViewCloseReason_|Required|**PpProtectedViewCloseReason**|A constant that specifies the reason the Protected View window is closed.|
| _Cancel_|Required|**Boolean**|**False** when the event occurs. If the event procedure sets this argument to **True**, the window does not close when the procedure is finished.|

## Return value

**Nothing**


## Remarks

If the  **ProtectedViewWindowsBeforeClose** event is called as part of the [ProtectedViewWindow.Edit](PowerPoint.ProtectedViewWindow.Edit.md) method, setting _Cancel_ to **True** produces no action.


## See also


[Application Object](PowerPoint.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]