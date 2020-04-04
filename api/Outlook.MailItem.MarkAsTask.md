---
title: MailItem.MarkAsTask method (Outlook)
keywords: vbaol11.chm3039
f1_keywords:
- vbaol11.chm3039
ms.prod: outlook
api_name:
- Outlook.MailItem.MarkAsTask
ms.assetid: ee38093d-a180-07f7-eae8-c9dbb2e8f413
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.MarkAsTask method (Outlook)

Marks a **[MailItem](Outlook.MailItem.md)** object as a task and assigns a task interval for the object.


## Syntax

_expression_. `MarkAsTask`( `_MarkInterval_` )

 _expression_ An expression that returns a [MailItem](Outlook.MailItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _MarkInterval_|Required| **[OlMarkInterval](Outlook.OlMarkInterval.md)**|The task interval for the  **MailItem**.|

## Remarks

Calling this method sets the value of several other properties, depending on the value provided in  _MarkInterval_. For more information about the properties set by specifying  _MarkInterval_, see [OlMarkInterval Enumeration](Outlook.OlMarkInterval.md).


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]