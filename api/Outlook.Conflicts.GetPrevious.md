---
title: Conflicts.GetPrevious method (Outlook)
keywords: vbaol11.chm409
f1_keywords:
- vbaol11.chm409
ms.prod: outlook
api_name:
- Outlook.Conflicts.GetPrevious
ms.assetid: 23b5d75a-e1eb-7164-df92-71e37a1ec79f
ms.date: 06/08/2017
localization_priority: Normal
---


# Conflicts.GetPrevious method (Outlook)

Returns the previous object in the  **[Conflicts](Outlook.Conflicts.md)** collection.


## Syntax

_expression_. `GetPrevious`

_expression_ A variable that represents a [Conflicts](Outlook.Conflicts.md) object.


## Return value

A **[Conflict](Outlook.Conflict.md)** object that represents the previous object contained by the collection.


## Remarks

It returns  **Nothing** if no previous object exists, for example, if already positioned at the beginning of the collection. To ensure correct operation of the **[GetFirst](Outlook.Conflicts.GetFirst.md)**, **[GetLast](Outlook.Conflicts.GetLast.md)**, **[GetNext](Outlook.Conflicts.GetNext.md)**, and **GetPrevious** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection, and call **GetLast** before calling **GetPrevious**. To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


[Conflicts Object](Outlook.Conflicts.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]