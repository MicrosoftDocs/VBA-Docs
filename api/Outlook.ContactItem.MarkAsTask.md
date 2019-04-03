---
title: ContactItem.MarkAsTask method (Outlook)
keywords: vbaol11.chm3031
f1_keywords:
- vbaol11.chm3031
ms.prod: outlook
api_name:
- Outlook.ContactItem.MarkAsTask
ms.assetid: def25d8d-6074-5e4d-18d9-82381b0b7876
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.MarkAsTask method (Outlook)

Marks a  **[ContactItem](Outlook.ContactItem.md)** object as a task and assigns a task interval for the object.


## Syntax

_expression_. `MarkAsTask`( `_MarkInterval_` )

 _expression_ An expression that returns a [ContactItem](Outlook.ContactItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _MarkInterval_|Required| **[OlMarkInterval](Outlook.OlMarkInterval.md)**|The task interval for the  **ContactItem**.|

## Remarks

Calling this method sets the value of several other properties, depending on the value provided in  _MarkInterval_. For more information about the properties set by specifying  _MarkInterval_, see [OlMarkInterval Enumeration](Outlook.OlMarkInterval.md).


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]