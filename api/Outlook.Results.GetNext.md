---
title: Results.GetNext method (Outlook)
keywords: vbaol11.chm507
f1_keywords:
- vbaol11.chm507
ms.prod: outlook
api_name:
- Outlook.Results.GetNext
ms.assetid: 3667738a-fcae-b786-e8d4-e478b1614c8c
ms.date: 06/08/2017
localization_priority: Normal
---


# Results.GetNext method (Outlook)

Returns the next object in the collection. 


## Syntax

_expression_. `GetNext`

_expression_ A variable that represents a [Results](Outlook.Results.md) object.


## Return value

An  **Object** value that represents the next object contained by the collection.


## Remarks

It returns  **Nothing** if no next object exists, for example, if already positioned at the end of the collection.To ensure correct operation of the **[GetFirst](Outlook.Results.GetFirst.md)**, **[GetLast](Outlook.Results.GetLast.md)**, **GetNext**, and **[GetPrevious](Outlook.Results.GetPrevious.md)** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection, and call **GetLast** before calling **GetPrevious**. To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


[Results Object](Outlook.Results.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]