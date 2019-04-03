---
title: MailItem.BeforeAutoSave event (Outlook)
ms.prod: outlook
api_name:
- Outlook.MailItem.BeforeAutoSave
ms.assetid: 0c725b91-f72f-7ceb-b2a9-da4f0369cf41
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.BeforeAutoSave event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

_expression_. `BeforeAutoSave`( `_Cancel_` , )

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[MailItem](Outlook.MailItem.md)** to be saved.|

## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]