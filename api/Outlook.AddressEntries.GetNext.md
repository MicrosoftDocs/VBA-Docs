---
title: AddressEntries.GetNext method (Outlook)
keywords: vbaol11.chm35
f1_keywords:
- vbaol11.chm35
ms.prod: outlook
api_name:
- Outlook.AddressEntries.GetNext
ms.assetid: 7579909c-90a2-660f-6cf5-039a441ccc93
ms.date: 06/08/2017
localization_priority: Normal
---


# AddressEntries.GetNext method (Outlook)

Returns the next object in the  **[AddressEntries](Outlook.AddressEntries.md)** collection.


## Syntax

_expression_. `GetNext`

_expression_ A variable that represents an [AddressEntries](Outlook.AddressEntries.md) object.


## Return value

An **[AddressEntry](Outlook.AddressEntry.md)** object that represents the next object contained by the collection.


## Remarks

It returns  **Nothing** if no next object exists, for example, if already positioned at the end of the collection.To ensure correct operation of the **[GetFirst](Outlook.AddressEntries.GetFirst.md)**, **[GetLast](Outlook.AddressEntries.GetLast.md)**, **GetNext**, and **[GetPrevious](Outlook.AddressEntries.GetPrevious.md)** methods in a large collection, call **GetFirst** before calling **GetNext** on that collection, and call **GetLast** before calling **GetPrevious**. To ensure that you are always making the calls on the same collection, create an explicit variable that refers to that collection before entering the loop.


## See also


[AddressEntries Object](Outlook.AddressEntries.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]