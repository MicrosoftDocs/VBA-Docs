---
title: Conflicts.GetLast method (Outlook)
keywords: vbaol11.chm407
f1_keywords:
- vbaol11.chm407
ms.prod: outlook
api_name:
- Outlook.Conflicts.GetLast
ms.assetid: 2f82fcab-7c8e-3df7-adc1-8f701d3bf9cb
ms.date: 06/08/2017
localization_priority: Normal
---


# Conflicts.GetLast method (Outlook)

Returns the last object in the  **[Conflicts](Outlook.Conflicts.md)** collection.


## Syntax

_expression_. `GetLast`

_expression_ A variable that represents a [Conflicts](Outlook.Conflicts.md) object.


## Return value

A  **[Conflict](Outlook.Conflict.md)** object that represents the last object contained by the collection.


## Remarks

 It returns **Nothing** if no last object exists, for example, if the collection is empty. To ensure correct operation of the **[GetFirst](Outlook.Conflicts.GetFirst.md)**, **GetLast**, **[GetNext](Outlook.Conflicts.GetNext.md)**, and **[GetPrevious](Outlook.Conflicts.GetPrevious.md)** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection, and call **GetLast** before calling **GetPrevious**. To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


[Conflicts Object](Outlook.Conflicts.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]