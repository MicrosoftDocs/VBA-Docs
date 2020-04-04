---
title: TimeZone.Bias property (Outlook)
keywords: vbaol11.chm3285
f1_keywords:
- vbaol11.chm3285
ms.prod: outlook
api_name:
- Outlook.TimeZone.Bias
ms.assetid: 18f55011-5d71-2e3b-4049-a37323f09478
ms.date: 06/08/2017
localization_priority: Normal
---


# TimeZone.Bias property (Outlook)

Returns a **Long** value that represents the difference in minutes of between the local time in this time zone and the Coordinated Universal Time (UTC). Read-only.


## Syntax

_expression_. `Bias`

_expression_ A variable that represents a [TimeZone](Outlook.TimeZone.md) object.


## Remarks

This value is stored as part of the value for  **TZI** for that time zone in the Windows registry. The **TZI** value is mapped to the Windows **[TIME_ZONE_INFORMATION](overview/Outlook.md)** structure.

 **Bias** does not take into account any time offset for daylight time or standard time in the time zone. To account for any daylight time offset, use **[DaylightBias](Outlook.TimeZone.DaylightBias.md)**. In general, when the local time zone is adopting daylight time, UTC time is the result of adding the **Bias** and **DaylightBias** to the local time. To account for any standard time offset, use **[StandardBias](Outlook.TimeZone.StandardBias.md)**. In general, when the local time zone is adopting standard time, UTC time is the result of adding the **Bias** and **StandardBias** to the local time.

For example, in a state adopting daylight time in the Pacific time zone, the  **Bias** is 480 minutes and **DaylightBias** is -60 minutes. To determine the time in UTC for June 11, 2 A.M. PST, add a **Bias** of (480/60) hours and a **DaylightBias** of -(60/60) hours to the local time June 11, 2 A.M. The time in UTC is June 11, 9 A.M.


## See also


[TimeZone Object](Outlook.TimeZone.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]