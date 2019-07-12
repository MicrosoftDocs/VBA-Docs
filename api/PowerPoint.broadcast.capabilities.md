---
title: Broadcast.Capabilities property (PowerPoint)
keywords: vbapp10.chm732011
f1_keywords:
- vbapp10.chm732011
ms.assetid: 9ac880c8-e222-08d7-1bea-cd9218d4c7f7
ms.date: 06/08/2017
ms.prod: powerpoint
localization_priority: Normal
---


# Broadcast.Capabilities property (PowerPoint)

Returns a  **Long** that represents the capabilities of the specified broadcast. Read-only.


## Syntax

_expression_. `Capabilities`

_expression_ A variable that represents a [Broadcast](PowerPoint.Broadcast.md) object.


## Remarks

The  **Capabilities** property can return the following[MSOBroadcastCapabilities](overview/Library-Reference/msobroadcastcapabilities-enumeration-office.md) values:



|Name|Value|Description|
|:-----|:-----|:-----|
|**MSOBroadcastCapFileSizeLimited**|1|File size limited|
|**MSOBroadcastCapSupportsMeetingNotes**|2|Supports meeting notes|
|**MSOBroadcastCapSupportsUpdateDoc**|4|Supports document update|

The values returned correspond to either Office or Microsoft Office 2010 broadcast presentation service capabilities.


## Property value

 **INT32**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]