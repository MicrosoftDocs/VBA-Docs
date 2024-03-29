---
title: OlkCheckBox.Exit event (Outlook)
keywords: vbaol11.chm1000155
f1_keywords:
- vbaol11.chm1000155
api_name:
- Outlook.OlkCheckBox.Exit
ms.assetid: a89b3d32-c540-ea72-b018-fabc9b9760f3
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# OlkCheckBox.Exit event (Outlook)

Occurs just after the focus passes from this control to another control on the same form.


## Syntax

_expression_. `Exit`( `_Cancel_` )

_expression_ A variable that represents an [OlkCheckBox](Outlook.OlkCheckBox.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the exit operation is not completed and the focus remains in this control.|

## See also


[OlkCheckBox Object](Outlook.OlkCheckBox.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]