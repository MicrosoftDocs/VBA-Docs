---
title: SharingItem.BeforeAutoSave event (Outlook)
ms.prod: outlook
api_name:
- Outlook.SharingItem.BeforeAutoSave
ms.assetid: 38515dda-2539-5f0b-4c04-831067c09327
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.BeforeAutoSave event (Outlook)

Occurs before the  **[SharingItem](Outlook.SharingItem.md)** is automatically saved by Outlook.


## Syntax

_expression_. `BeforeAutoSave`( `_Cancel_` )

 _expression_ An expression that returns a [SharingItem](Outlook.SharingItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **SharingItem** to be saved.|

## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]