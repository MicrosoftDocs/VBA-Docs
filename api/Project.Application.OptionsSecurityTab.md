---
title: Application.OptionsSecurityTab method (Project)
keywords: vbapj.chm2504
f1_keywords:
- vbapj.chm2504
ms.prod: project-server
api_name:
- Project.Application.OptionsSecurityTab
ms.assetid: f19ecd9c-2507-e437-7780-cf4998b7fd48
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.OptionsSecurityTab method (Project)

Displays a specific tab of the **Trust Center** dialog box in Project.


## Syntax

_expression_. `OptionsSecurityTab`( `_DefaultTab_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DefaultTab_|Optional|**PjOptionsSecurityTab**|Specifies the tab to open in the **Trust Center** dialog box. Can be one of the **[PjOptionsSecurityTab](Project.PjOptionsSecurityTab.md)** constants. The default is **pjOptionsSecurityTabPublishers** for the **Trusted Publishers** tab.|

## Return value

 **Boolean**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]