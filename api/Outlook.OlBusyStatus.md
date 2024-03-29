---
title: OlBusyStatus enumeration (Outlook)
keywords: vbaol11.chm3053
f1_keywords:
- vbaol11.chm3053
api_name:
- Outlook.OlBusyStatus
ms.assetid: 4391ccb4-a035-30d1-9693-61b83050b31f
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# OlBusyStatus enumeration (Outlook)

Indicates a user's availability.



|Name|Value|Description|
|:-----|:-----|:-----|
| **olBusy**|2|The user is busy.|
| **olFree**|0|The user is available.|
| **olOutOfOffice**|3|The user is out of office.|
| **olTentative**|1|The user has a tentative appointment scheduled.|
| **olWorkingElsewhere**|4|The user is working in a location away from the office.|

## Remarks

The user's availability is based on scheduled appointments. See [AppointmentItem.BusyStatus property (Outlook)](Outlook.AppointmentItem.BusyStatus.md).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]