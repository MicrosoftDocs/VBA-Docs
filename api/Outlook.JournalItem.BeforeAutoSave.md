---
title: JournalItem.BeforeAutoSave event (Outlook)
ms.prod: outlook
api_name:
- Outlook.JournalItem.BeforeAutoSave
ms.assetid: b4924fd8-52cd-fa8d-11d8-2683ea2f5b52
ms.date: 06/08/2017
localization_priority: Normal
---


# JournalItem.BeforeAutoSave event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

_expression_. `BeforeAutoSave`( `_Cancel_` , )

_expression_ A variable that represents a [JournalItem](Outlook.JournalItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[JournalItem](Outlook.JournalItem.md)** to be saved.|

## See also


[JournalItem Object](Outlook.JournalItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]