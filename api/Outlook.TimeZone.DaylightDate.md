---
title: TimeZone.DaylightDate property (Outlook)
keywords: vbaol11.chm3289
f1_keywords:
- vbaol11.chm3289
ms.prod: outlook
api_name:
- Outlook.TimeZone.DaylightDate
ms.assetid: a653b0ec-1462-165f-36e3-1be57513a2c7
ms.date: 06/08/2017
localization_priority: Normal
---


# TimeZone.DaylightDate property (Outlook)

Returns a  **Date** value that represents the date and time in this time zone when time changes over to daylight time in the current year. Read-only.


## Syntax

_expression_. `DaylightDate`

_expression_ A variable that represents a [TimeZone](Outlook.TimeZone.md) object.


## Remarks

This value is stored as part of the  **TZI** value for the time zone in the Windows registry. The **TZI** value is mapped to the Windows **[TIME_ZONE_INFORMATION](overview/Outlook.md)** structure.


## See also


[TimeZone Object](Outlook.TimeZone.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]