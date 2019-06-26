---
title: MeetingItem.ReadComplete event (Outlook)
ms.assetid: 17ef8085-38ac-7e32-7704-54a2f2224e87
ms.date: 06/08/2017
ms.prod: outlook
localization_priority: Normal
---


# MeetingItem.ReadComplete event (Outlook)
Occurs when Outlook has completed reading the properties of the item.

## Version information

Version Added: Outlook 2013 


## Syntax

_expression_. `ReadComplete`_(Cancel)_

_expression_ A variable that represents a [MeetingItem](Outlook.MeetingItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the read operation is not completed and the item is not displayed in the Reading Pane or inspector.|

## Remarks

The  **ReadComplete** event occurs after the [BeforeRead](Outlook.MeetingItem.BeforeRead.md) event and before the [Read](Outlook.MeetingItem.Read.md) event for the item.

To determine when the item is unloaded from memory, use the [Unload](Outlook.MeetingItem.Unload.md) event.

The  **ReadComplete** event corresponds to the Exchange Client Extensions (ECE) event **IExchExtMessageEvents::OnReadComplete**.


## See also


[MeetingItem Object](Outlook.MeetingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]