---
title: Results.GetPrevious method (Outlook)
keywords: vbaol11.chm508
f1_keywords:
- vbaol11.chm508
ms.prod: outlook
api_name:
- Outlook.Results.GetPrevious
ms.assetid: be9877c4-602d-7e2d-a00b-edb4aead7441
ms.date: 06/08/2017
localization_priority: Normal
---


# Results.GetPrevious method (Outlook)

Returns the previous object in the collection. 


## Syntax

_expression_. `GetPrevious`

_expression_ A variable that represents a [Results](Outlook.Results.md) object.


## Return value

An  **Object** value that represents the previous object contained by the collection.


## Remarks

It returns  **Nothing** if no previous object exists, for example, if already positioned at the beginning of the collection.To ensure correct operation of the **[GetFirst](Outlook.Results.GetFirst.md)**, **[GetLast](Outlook.Results.GetLast.md)**, **[GetNext](Outlook.Results.GetNext.md)**, and **GetPrevious** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection, and call **GetLast** before calling **GetPrevious**. To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


[Results Object](Outlook.Results.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]