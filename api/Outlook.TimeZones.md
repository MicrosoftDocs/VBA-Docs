---
title: TimeZones object (Outlook)
keywords: vbaol11.chm3300
f1_keywords:
- vbaol11.chm3300
ms.prod: outlook
api_name:
- Outlook.TimeZones
ms.assetid: c68f8589-44e9-3c12-45c1-96943fa9bcb7
ms.date: 06/08/2017
localization_priority: Normal
---


# TimeZones object (Outlook)

A collection of  **[TimeZone](Outlook.TimeZone.md)** objects.


## Remarks

This collection is read-only, and serves the purpose of enumerating time zones supported by Windows and thus Outlook. Its value is accessible through the property  **[Application.TimeZones](Outlook.Application.TimeZones.md)** and is based on the data stored in the Windows registry key HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Time Zones.


## Methods



|Name|
|:-----|
|[ConvertTime](Outlook.TimeZones.ConvertTime.md)|
|[Item](Outlook.TimeZones.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Outlook.TimeZones.Application.md)|
|[Class](Outlook.TimeZones.Class.md)|
|[Count](Outlook.TimeZones.Count.md)|
|[CurrentTimeZone](Outlook.TimeZones.CurrentTimeZone.md)|
|[Parent](Outlook.TimeZones.Parent.md)|
|[Session](Outlook.TimeZones.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]