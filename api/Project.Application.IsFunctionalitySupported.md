---
title: Application.IsFunctionalitySupported event (Project)
ms.prod: project-server
api_name:
- Project.Application.IsFunctionalitySupported
ms.assetid: f6462a3b-5a36-3b2e-79bd-78cce567aed8
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.IsFunctionalitySupported event (Project)

Occurs after the **LoadWebBrowserControl** method is called with the third parameter ( _FunctionalityName_) set.


## Syntax

_expression_. `IsFunctionalitySupported`( `_bstrFunctionality_`, `_Info_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _bstrFunctionality_|Required|**String**|Name of the functionality.|
| _Info_|Required|**EventInfo**|EventInfo object.|

## Return value

**Nothing**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]