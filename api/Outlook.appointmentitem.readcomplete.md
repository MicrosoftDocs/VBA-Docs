---
title: AppointmentItem.ReadComplete event (Outlook)
ms.assetid: 749e8d58-c15c-0b63-5486-cc2aa2190638
ms.date: 06/08/2017
ms.prod: outlook
localization_priority: Normal
---


# AppointmentItem.ReadComplete event (Outlook)
Occurs when Outlook has completed reading the properties of the item.

## Syntax

_expression_. `ReadComplete`_(Cancel)_

_expression_ A variable that represents an [AppointmentItem](Outlook.AppointmentItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the read operation is not completed and the item is not displayed in the Reading Pane or inspector.|

## Remarks

The **ReadComplete** event occurs after the [BeforeRead](Outlook.AppointmentItem.BeforeRead.md) event and before the [Read](Outlook.AppointmentItem.Read.md) event for the item.

To determine when the item is unloaded from memory, use the [Unload](Outlook.AppointmentItem.Unload.md) event.

The **ReadComplete** event corresponds to the Exchange Client Extensions (ECE) event **IExchExtMessageEvents::OnReadComplete**.


## See also


[AppointmentItem Object](Outlook.AppointmentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]