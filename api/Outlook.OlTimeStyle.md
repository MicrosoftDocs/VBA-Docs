---
title: OlTimeStyle enumeration (Outlook)
keywords: vbaol11.chm1000031
f1_keywords:
- vbaol11.chm1000031
ms.prod: outlook
api_name:
- Outlook.OlTimeStyle
ms.assetid: 82c4d063-29f2-d7c8-44ff-8b4aca912855
ms.date: 06/08/2017
localization_priority: Normal
---


# OlTimeStyle enumeration (Outlook)

Specifies how time values are displayed and how entries of time values are interpreted by a time control.



|Name|Value|Description|
|:-----|:-----|:-----|
| **olTimeStyleShortDuration**|4|The drop-down portion of the time control displays only duration values with the interval set by the  **[OlkTimeControl.IntervalTime](Outlook.OlkTimeControl.IntervalTime.md)** property.|
| **olTimeStyleTimeDuration**|1|The drop-down portion of the time control displays time values starting from the  **[ReferenceTime](Outlook.OlkTimeControl.ReferenceTime.md)** and uses the **OlkTimeControl.IntervalTime** property as the increment. The edit box of the time control displays the duration from the **ReferenceTime** to the selected time.|
| **olTimeStyleTimeOnly**|0|The drop-down portion of the time control displays only time values with the interval set by the  **OlkTimeControl.IntervalTime** property.|

## Remarks

Use the time control with the  **olTimeStyleShortDuration** style for duration fields, such as the [Duration](Outlook.JournalItem.Duration.md) of a [JournalItem](Outlook.JournalItem.md). Use the time control with the  **olTimeStyleTimeDuration** style for the end time of an appointment item. Use the time control with the **olTimeStyleTimeOnly** style for the start time of an appointment item.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]