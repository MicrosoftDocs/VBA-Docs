---
title: RemoteItem.BeforeAutoSave event (Outlook)
ms.prod: outlook
api_name:
- Outlook.RemoteItem.BeforeAutoSave
ms.assetid: f33e1442-0e65-cc78-34ac-496b65ba565e
ms.date: 06/08/2017
localization_priority: Normal
---


# RemoteItem.BeforeAutoSave event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

_expression_. `BeforeAutoSave`( `_Cancel_` )

_expression_ A variable that represents a [RemoteItem](Outlook.RemoteItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[RemoteItem](Outlook.RemoteItem.md)** to be saved.|

## See also


[RemoteItem Object](Outlook.RemoteItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]