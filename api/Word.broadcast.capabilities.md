---
title: Broadcast.Capabilities property (Word)
keywords: vbawd10.chm36438019
f1_keywords:
- vbawd10.chm36438019
ms.prod: word
ms.assetid: 86388adc-95c3-3c06-dbfe-a0455e93c90f
ms.date: 06/08/2017
localization_priority: Normal
---


# Broadcast.Capabilities property (Word)

Returns a  **Long** that represents the capabilities of the specified broadcast. Read-only.


## Syntax

_expression_. `Capabilities`

_expression_ A variable that represents a **[Broadcast](Word.broadcast.md)** object.


## Remarks

The  **Capabilities** property can return the following[MSOBroadcastCapabilities](overview/Library-Reference/msobroadcastcapabilities-enumeration-office.md) values:



|Name|Value|Description|
|:-----|:-----|:-----|
| **MSOBroadcastCapFileSizeLimited**|1|File size limited|
| **MSOBroadcastCapSupportsMeetingNotes**|2|Supports meeting notes|
| **MSOBroadcastCapSupportsUpdateDoc**|4|Supports document update|

The values returned correspond to either Office or Microsoft Office 2010 broadcast presentation service capabilities.


## Property value

 **INT32**


## See also


[Broadcast Object](Word.broadcast.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]