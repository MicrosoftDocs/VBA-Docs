---
title: Inspector.PageChange event (Outlook)
keywords: vbaol11.chm472
f1_keywords:
- vbaol11.chm472
ms.prod: outlook
api_name:
- Outlook.Inspector.PageChange
ms.assetid: f0ba9820-84bf-2367-364a-519e6ed88289
ms.date: 06/08/2017
localization_priority: Normal
---


# Inspector.PageChange event (Outlook)

Occurs when the active form page changes, either programmatically or by user action, on an [Inspector](Outlook.Inspector.md) object.


## Syntax

_expression_. `PageChange`( `_ActivePageName_` )

_expression_ A variable that represents an [Inspector](Outlook.Inspector.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ActivePageName_|Required| **String**|The name of the active page.|

## Remarks

An error occurs if the event handler for this event calls either the  **[Close](Outlook.Inspector.Close(method).md)** or **[SetCurrentFormPage](Outlook.Inspector.SetCurrentFormPage.md)** methods.


## See also


[Inspector Object](Outlook.Inspector.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]