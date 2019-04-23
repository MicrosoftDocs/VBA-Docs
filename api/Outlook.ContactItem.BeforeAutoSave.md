---
title: ContactItem.BeforeAutoSave event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ContactItem.BeforeAutoSave
ms.assetid: c9fe9c4d-3c00-455c-3e89-9ac584597117
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.BeforeAutoSave event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

_expression_. `BeforeAutoSave`( `_Cancel_` , )

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[ContactItem](Outlook.ContactItem.md)** to be saved.|

## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]