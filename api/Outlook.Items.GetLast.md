---
title: Items.GetLast method (Outlook)
keywords: vbaol11.chm65
f1_keywords:
- vbaol11.chm65
ms.prod: outlook
api_name:
- Outlook.Items.GetLast
ms.assetid: d02a20be-19fc-fb6e-feff-b66ca0273beb
ms.date: 06/08/2017
localization_priority: Normal
---


# Items.GetLast method (Outlook)

Returns the last object in the collection. 


## Syntax

_expression_. `GetLast`

_expression_ A variable that represents an [Items](Outlook.Items.md) object.


## Return value

An  **Object** value that represents the last object contained by the collection.


## Remarks

It returns  **Nothing** if no last object exists, for example, if the collection is empty. To ensure correct operation of the **[GetFirst](Outlook.Items.GetFirst.md)**, **GetLast**, **[GetNext](Outlook.Items.GetNext.md)**, and **[GetPrevious](Outlook.Items.GetPrevious.md)** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection, and call **GetLast** before calling **GetPrevious**. To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


[Items Object](Outlook.Items.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]