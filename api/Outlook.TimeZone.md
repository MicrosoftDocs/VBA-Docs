---
title: TimeZone object (Outlook)
keywords: vbaol11.chm3299
f1_keywords:
- vbaol11.chm3299
ms.prod: outlook
api_name:
- Outlook.TimeZone
ms.assetid: b27da70d-e545-cc13-9529-cfd327ab7a7c
ms.date: 06/08/2017
localization_priority: Normal
---


# TimeZone object (Outlook)

Represents information for a time zone as supported by Windows.


## Remarks

The **TimeZone** object is an Outlook wrapper for time zone data.

This data can be obtained from the Windows registry key HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Time Zones. In this case, some properties of this object are parts of in the  **TZI** value for the time zone in the registry. A **TZI** value is mapped to the Windows **[TIME_ZONE_INFORMATION](overview/Outlook.md)** structure.


## Properties



|Name|
|:-----|
|[Application](Outlook.TimeZone.Application.md)|
|[Bias](Outlook.TimeZone.Bias.md)|
|[Class](Outlook.TimeZone.Class.md)|
|[DaylightBias](Outlook.TimeZone.DaylightBias.md)|
|[DaylightDate](Outlook.TimeZone.DaylightDate.md)|
|[DaylightDesignation](Outlook.TimeZone.DaylightDesignation.md)|
|[ID](Outlook.TimeZone.ID.md)|
|[Name](Outlook.TimeZone.Name.md)|
|[Parent](Outlook.TimeZone.Parent.md)|
|[Session](Outlook.TimeZone.Session.md)|
|[StandardBias](Outlook.TimeZone.StandardBias.md)|
|[StandardDate](Outlook.TimeZone.StandardDate.md)|
|[StandardDesignation](Outlook.TimeZone.StandardDesignation.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]