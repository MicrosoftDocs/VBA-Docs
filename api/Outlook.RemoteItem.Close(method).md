---
title: RemoteItem.Close method (Outlook)
keywords: vbaol11.chm1612
f1_keywords:
- vbaol11.chm1612
ms.prod: outlook
api_name:
- Outlook.RemoteItem.Close
ms.assetid: 274e73b2-d5bf-1add-6add-e9d571f14d2a
ms.date: 06/08/2017
localization_priority: Normal
---


# RemoteItem.Close method (Outlook)

Closes and optionally saves changes to the Outlook item.


## Syntax

_expression_.**Close** (_SaveMode_)

_expression_ A variable that represents a '[RemoteItem](Outlook.RemoteItem.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SaveMode_|Required| **[OlInspectorClose](Outlook.OlInspectorClose.md)**|The close behavior. If the item displayed within the inspector has not been changed, this argument has no effect.|

## See also


[RemoteItem Object](Outlook.RemoteItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]