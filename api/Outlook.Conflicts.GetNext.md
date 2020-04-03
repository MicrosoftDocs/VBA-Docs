---
title: Conflicts.GetNext method (Outlook)
keywords: vbaol11.chm408
f1_keywords:
- vbaol11.chm408
ms.prod: outlook
api_name:
- Outlook.Conflicts.GetNext
ms.assetid: 2e21ea88-c732-17ee-cd87-698fee992269
ms.date: 06/08/2017
localization_priority: Normal
---


# Conflicts.GetNext method (Outlook)

Returns the next object in the  **[Conflicts](Outlook.Conflicts.md)** collection.


## Syntax

_expression_. `GetNext`

_expression_ A variable that represents a [Conflicts](Outlook.Conflicts.md) object.


## Return value

A **[Conflict](Outlook.Conflict.md)** object that represents the next object contained by the collection.


## Remarks

It returns  **Nothing** if no next object exists, for example, if already positioned at the end of the collection. To ensure correct operation of the **[GetFirst](Outlook.Conflicts.GetFirst.md)**, **[GetLast](Outlook.Conflicts.GetLast.md)**, **GetNext**, and **[GetPrevious](Outlook.Conflicts.GetPrevious.md)** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection, and call **GetLast** before calling **GetPrevious**. To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


[Conflicts Object](Outlook.Conflicts.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]