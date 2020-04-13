---
title: ExchangeUser.GetDirectReports method (Outlook)
keywords: vbaol11.chm2083
f1_keywords:
- vbaol11.chm2083
ms.prod: outlook
api_name:
- Outlook.ExchangeUser.GetDirectReports
ms.assetid: 753201ad-8001-3185-7d68-fda15907099d
ms.date: 06/08/2017
localization_priority: Normal
---


# ExchangeUser.GetDirectReports method (Outlook)

Obtains an **[AddressEntries](Outlook.AddressEntries.md)** collection object that contains all the users directly reporting to the Exchange user.


## Syntax

_expression_. `GetDirectReports`

_expression_ A variable that represents an [ExchangeUser](Outlook.ExchangeUser.md) object.


## Return value

An **AddressEntries** collection object that contains the users directly reporting to the Exchange user. The **AddressEntries** object will have a count of zero (0) if there is no direct report represented by an **[AddressEntry](Outlook.AddressEntry.md)** in the current session, or if direct reports have not been implemented in the Exchange directory.


## Remarks

 **GetDirectReports** is an expensive operation in terms of performance if there is a slow connection to the Exchange server.


## See also


[ExchangeUser Object](Outlook.ExchangeUser.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]