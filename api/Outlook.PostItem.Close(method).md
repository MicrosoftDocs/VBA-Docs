---
title: PostItem.Close method (Outlook)
keywords: vbaol11.chm1539
f1_keywords:
- vbaol11.chm1539
ms.prod: outlook
api_name:
- Outlook.PostItem.Close
ms.assetid: fd80ee3c-2ee1-20ff-1f43-d706695b128c
ms.date: 06/08/2017
localization_priority: Normal
---


# PostItem.Close method (Outlook)

Closes and optionally saves changes to the Outlook item.


## Syntax

_expression_.**Close** (_SaveMode_)

_expression_ A variable that represents a [PostItem](Outlook.PostItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SaveMode_|Required| **[OlInspectorClose](Outlook.OlInspectorClose.md)**|The close behavior. If the item displayed within the inspector has not been changed, this argument has no effect.|

## See also


[PostItem Object](Outlook.PostItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]