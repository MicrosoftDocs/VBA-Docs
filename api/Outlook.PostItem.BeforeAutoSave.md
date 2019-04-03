---
title: PostItem.BeforeAutoSave event (Outlook)
ms.prod: outlook
api_name:
- Outlook.PostItem.BeforeAutoSave
ms.assetid: 61a44326-0215-869b-0824-2308fd8017cf
ms.date: 06/08/2017
localization_priority: Normal
---


# PostItem.BeforeAutoSave event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

_expression_. `BeforeAutoSave`( `_Cancel_` )

_expression_ A variable that represents a [PostItem](Outlook.PostItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[PostItem](Outlook.PostItem.md)** to be saved.|

## See also


[PostItem Object](Outlook.PostItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]