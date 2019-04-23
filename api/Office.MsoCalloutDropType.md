---
title: MsoCalloutDropType enumeration (Office)
ms.prod: office
api_name:
- Office.MsoCalloutDropType
ms.assetid: 0923e0a7-beb6-224f-6a87-85111f58ae3b
ms.date: 01/31/2019
localization_priority: Normal
---


# MsoCalloutDropType enumeration (Office)

Specifies the starting position of the callout line relative to the text bounding box. Used with the **PresetDrop** method of the **CalloutFormat** object.

<br/>

|Name|Value|Description|
|:-----|:-----|:-----|
|**msoCalloutDropBottom**|4|Bottom|
|**msoCalloutDropCenter**|3|Center|
|**msoCalloutDropCustom**|1|Custom. If this value is used as the value for the **PresetDrop** property, the **Drop** and **AutoAttach** properties of the **CalloutFormat** object are used to determine where the callout line attaches to the text box.|
|**msoCalloutDropMixed**|-2|Return value only; indicates a combination of the other states. |
|**msoCalloutDropTop**|2|Top|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]