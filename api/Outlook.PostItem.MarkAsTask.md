---
title: PostItem.MarkAsTask method (Outlook)
keywords: vbaol11.chm3043
f1_keywords:
- vbaol11.chm3043
ms.prod: outlook
api_name:
- Outlook.PostItem.MarkAsTask
ms.assetid: 78ead34b-3861-0204-1bc3-687a2c25ab73
ms.date: 06/08/2017
localization_priority: Normal
---


# PostItem.MarkAsTask method (Outlook)

Marks a **[PostItem](Outlook.PostItem.md)** object as a task and assigns a task interval for the object.


## Syntax

_expression_. `MarkAsTask`( `_MarkInterval_` )

 _expression_ An expression that returns a [PostItem](Outlook.PostItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _MarkInterval_|Required| **[OlMarkInterval](Outlook.OlMarkInterval.md)**|The task interval for the  **PostItem**.|

## Remarks

Calling this method sets the value of several other properties, depending on the value provided in  _MarkInterval_. For more information about the properties set by specifying  _MarkInterval_, see [OlMarkInterval Enumeration](Outlook.OlMarkInterval.md).


## See also


[PostItem Object](Outlook.PostItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]