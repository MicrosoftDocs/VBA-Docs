---
title: OlkTimeControl.ReferenceTime property (Outlook)
keywords: vbaol11.chm1000391
f1_keywords:
- vbaol11.chm1000391
ms.prod: outlook
api_name:
- Outlook.OlkTimeControl.ReferenceTime
ms.assetid: 3979de6d-4992-f42c-b894-7f9661826ca6
ms.date: 06/08/2017
localization_priority: Normal
---


# OlkTimeControl.ReferenceTime property (Outlook)

Returns or sets a **Date** that specifies a reference time used for the **olTimeStyleTimeDuration** style setting on the time control. Read/write.


## Syntax

_expression_. `ReferenceTime`

_expression_ A variable that represents an [OlkTimeControl](Outlook.OlkTimeControl.md) object.


## Remarks

The default value is 12/30/1899.

When  **[Style](Outlook.OlkTimeControl.Style.md)** is **olTimeStyleTimeDuration**, the date control displays the value of **ReferenceTime** as the first selectable time value, shows additional intervals (specified by **[IntervalTime](Outlook.OlkTimeControl.IntervalTime.md)**) starting from the **ReferenceTime** value, and shows the duration of an event.

The default value for  **ReferenceTime** is 30 Dec 1899 12:00 AM. In this case, the time control will display **30 Dec 1899 12:00 AM** as the first selectable time. A value of 60 for **IntervalTime** will mark **30 Dec 1899 1:00 AM** as the first interval.


## See also


[OlkTimeControl Object](Outlook.OlkTimeControl.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]