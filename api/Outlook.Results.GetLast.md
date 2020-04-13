---
title: Results.GetLast method (Outlook)
keywords: vbaol11.chm506
f1_keywords:
- vbaol11.chm506
ms.prod: outlook
api_name:
- Outlook.Results.GetLast
ms.assetid: 90a50739-b9a9-92de-516b-1cd9f3fe8d50
ms.date: 06/08/2017
localization_priority: Normal
---


# Results.GetLast method (Outlook)

Returns the last object in the collection. 


## Syntax

_expression_. `GetLast`

_expression_ A variable that represents a [Results](Outlook.Results.md) object.


## Return value

An **Object** value that represents the last object contained by the collection.


## Remarks

It returns  **Nothing** if no last object exists, for example, if the collection is empty. To ensure correct operation of the **[GetFirst](Outlook.Results.GetFirst.md)**, **GetLast**, **[GetNext](Outlook.Results.GetNext.md)**, and **[GetPrevious](Outlook.Results.GetPrevious.md)** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection, and call **GetLast** before calling **GetPrevious**. To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


[Results Object](Outlook.Results.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]