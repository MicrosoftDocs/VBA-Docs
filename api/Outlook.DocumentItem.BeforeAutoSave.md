---
title: DocumentItem.BeforeAutoSave event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DocumentItem.BeforeAutoSave
ms.assetid: 3aaf57a3-bcc2-d0ba-6fd9-d801452dc4ca
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentItem.BeforeAutoSave event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

_expression_. `BeforeAutoSave`( `_Cancel_` )

_expression_ A variable that represents a [DocumentItem](Outlook.DocumentItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[DocumentItem](Outlook.DocumentItem.md)** to be saved.|

## See also


[DocumentItem Object](Outlook.DocumentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]