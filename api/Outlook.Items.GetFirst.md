---
title: Items.GetFirst method (Outlook)
keywords: vbaol11.chm64
f1_keywords:
- vbaol11.chm64
ms.prod: outlook
api_name:
- Outlook.Items.GetFirst
ms.assetid: 142a6174-118e-6256-0511-8ae9e142e555
ms.date: 06/08/2017
localization_priority: Normal
---


# Items.GetFirst method (Outlook)

Returns the first object in the collection. 


## Syntax

_expression_. `GetFirst`

_expression_ A variable that represents an [Items](Outlook.Items.md) object.


## Return value

An **Object** value that represents the first object contained by the collection.


## Remarks

Returns  **Nothing** if no first object exists, for example, if there are no objects in the collection. To ensure correct operation of the **GetFirst**, **[GetLast](Outlook.Items.GetLast.md)**, **[GetNext](Outlook.Items.GetNext.md)**, and **[GetPrevious](Outlook.Items.GetPrevious.md)** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection and call **GetLast** before calling **GetPrevious**. To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


[Items Object](Outlook.Items.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]