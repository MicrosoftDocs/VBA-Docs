---
title: ReportItem.BeforeAutoSave event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ReportItem.BeforeAutoSave
ms.assetid: c3a2882c-ff82-39a1-3d18-5bf4f608b09e
ms.date: 06/08/2017
localization_priority: Normal
---


# ReportItem.BeforeAutoSave event (Outlook)

Occurs before the item is automatically saved by Outlook.


## Syntax

_expression_. `BeforeAutoSave`( `_Cancel_` )

_expression_ A variable that represents a [ReportItem](Outlook.ReportItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|Set to  **True** to cancel the operation; otherwise, set to **False** to allow the **[ReportItem](Outlook.ReportItem.md)** to be saved.|

## See also


[ReportItem Object](Outlook.ReportItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]