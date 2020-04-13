---
title: Application.Alerts method (Project)
keywords: vbapj.chm10
f1_keywords:
- vbapj.chm10
ms.prod: project-server
api_name:
- Project.Application.Alerts
ms.assetid: 58c935d9-35a3-953b-4003-dc88f8532854
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Alerts method (Project)

Determines whether alerts appear when a macro runs.


## Syntax

_expression_. `Alerts`( `_Show_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Show_|Optional|**Boolean**|**True** if Project displays error messages when a macro runs. The default value is **True**.|

## Return value

 **Boolean**


## Remarks

The **Alerts** method applies only to the macro that contains the method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]