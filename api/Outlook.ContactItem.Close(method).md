---
title: ContactItem.Close method (Outlook)
keywords: vbaol11.chm956
f1_keywords:
- vbaol11.chm956
ms.prod: outlook
api_name:
- Outlook.ContactItem.Close
ms.assetid: 17cd04b5-1bf1-5df1-b1f4-f6e488d00fd5
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.Close method (Outlook)

Closes and optionally saves changes to the Outlook item.


## Syntax

_expression_.**Close** (_SaveMode_)

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SaveMode_|Required| **[OlInspectorClose](Outlook.OlInspectorClose.md)**|The close behavior. If the item displayed within the inspector has not been changed, this argument has no effect.|

## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]