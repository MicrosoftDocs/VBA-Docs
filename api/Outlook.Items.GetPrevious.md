---
title: Items.GetPrevious method (Outlook)
keywords: vbaol11.chm67
f1_keywords:
- vbaol11.chm67
ms.prod: outlook
api_name:
- Outlook.Items.GetPrevious
ms.assetid: 5dde47f8-2bd8-fdbe-d6e7-b1381e8a97a6
ms.date: 06/08/2017
localization_priority: Normal
---


# Items.GetPrevious method (Outlook)

Returns the previous object in the collection. 


## Syntax

_expression_. `GetPrevious`

_expression_ A variable that represents an [Items](Outlook.Items.md) object.


## Return value

An  **Object** value that represents the previous object contained by the collection.


## Remarks

It returns  **Nothing** if no previous object exists, for example, if already positioned at the beginning of the collection. To ensure correct operation of the **[GetFirst](Outlook.Items.GetFirst.md)**, **[GetLast](Outlook.Items.GetLast.md)**, **[GetNext](Outlook.Items.GetNext.md)**, and **GetPrevious** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection, and call **GetLast** before calling **GetPrevious**. To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


[Items Object](Outlook.Items.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]