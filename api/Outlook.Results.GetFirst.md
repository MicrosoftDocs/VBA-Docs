---
title: Results.GetFirst method (Outlook)
keywords: vbaol11.chm505
f1_keywords:
- vbaol11.chm505
ms.prod: outlook
api_name:
- Outlook.Results.GetFirst
ms.assetid: 9a8b56ce-5e93-f1b1-be7f-7734d86f4997
ms.date: 06/08/2017
localization_priority: Normal
---


# Results.GetFirst method (Outlook)

Returns the first object in the collection.


## Syntax

_expression_. `GetFirst`

_expression_ A variable that represents a [Results](Outlook.Results.md) object.


## Return value

An  **Object** value that represents the first object contained by the collection.


## Remarks

Returns  **Nothing** if no first object exists, for example, if there are no objects in the collection. To ensure correct operation of the **GetFirst**, **[GetLast](Outlook.Results.GetLast.md)**, **[GetNext](Outlook.Results.GetNext.md)**, and **[GetPrevious](Outlook.Results.GetPrevious.md)** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection and call **GetLast** before calling **GetPrevious**. To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


[Results Object](Outlook.Results.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]