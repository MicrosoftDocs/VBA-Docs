---
title: AddressEntries.GetFirst method (Outlook)
keywords: vbaol11.chm33
f1_keywords:
- vbaol11.chm33
ms.prod: outlook
api_name:
- Outlook.AddressEntries.GetFirst
ms.assetid: f8f03b6e-d79e-09b5-2f75-6886e699a4b3
ms.date: 06/08/2017
localization_priority: Normal
---


# AddressEntries.GetFirst method (Outlook)

Returns the first object in the  **[AddressEntries](Outlook.AddressEntries.md)** collection.


## Syntax

_expression_. `GetFirst`

_expression_ A variable that represents an [AddressEntries](Outlook.AddressEntries.md) object.


## Return value

An  **[AddressEntry](Outlook.AddressEntry.md)** object that represents the first object contained by the collection.


## Remarks

Returns  **Nothing** if no first object exists, for example, if there are no objects in the collection. To ensure correct operation of the **GetFirst**, **[GetLast](Outlook.AddressEntries.GetLast.md)**, **[GetNext](Outlook.AddressEntries.GetNext.md)**, and **[GetPrevious](Outlook.AddressEntries.GetPrevious.md)** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection and call **GetLast** before calling **GetPrevious**. To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


[AddressEntries Object](Outlook.AddressEntries.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]