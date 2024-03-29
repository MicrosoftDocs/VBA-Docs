---
title: Inspector.BeforeMaximize event (Outlook)
keywords: vbaol11.chm468
f1_keywords:
- vbaol11.chm468
api_name:
- Outlook.Inspector.BeforeMaximize
ms.assetid: 9793d228-85ea-50cd-4c1b-74ca23788aad
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Inspector.BeforeMaximize event (Outlook)

Occurs when an inspector is maximized by the user.


## Syntax

_expression_. `BeforeMaximize`( `_Cancel_` )

_expression_ A variable that represents an [Inspector](Outlook.Inspector.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the operation is not completed and the explorer or inspector is not maximized.|

## Remarks

This event can be cancelled after it has started.


## See also


[Inspector Object](Outlook.Inspector.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]