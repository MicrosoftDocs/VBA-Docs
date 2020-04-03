---
title: Conflicts.GetFirst method (Outlook)
keywords: vbaol11.chm406
f1_keywords:
- vbaol11.chm406
ms.prod: outlook
api_name:
- Outlook.Conflicts.GetFirst
ms.assetid: f257a9f1-d9ec-c13a-62f7-0228d55342da
ms.date: 06/08/2017
localization_priority: Normal
---


# Conflicts.GetFirst method (Outlook)

Returns the first object in the  **[Conflicts](Outlook.Conflicts.md)** collection.


## Syntax

_expression_. `GetFirst`

_expression_ A variable that represents a [Conflicts](Outlook.Conflicts.md) object.


## Return value

A  **[Conflict](Outlook.Conflict.md)** object that represents the first object contained by the collection.


## Remarks

Returns  **Nothing** if no first object exists, for example, if there are no objects in the collection. To ensure correct operation of the **GetFirst**, **[GetLast](Outlook.Conflicts.GetLast.md)**, **[GetNext](Outlook.Conflicts.GetNext.md)**, and **[GetPrevious](Outlook.Conflicts.GetPrevious.md)** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection and call **GetLast** before calling **GetPrevious**. To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


[Conflicts Object](Outlook.Conflicts.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]