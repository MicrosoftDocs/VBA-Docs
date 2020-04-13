---
title: TaskRequestAcceptItem.ReadComplete event (Outlook)
ms.assetid: 95718369-d2f8-31b9-145a-f53f242c0bfa
ms.date: 06/08/2017
ms.prod: outlook
localization_priority: Normal
---


# TaskRequestAcceptItem.ReadComplete event (Outlook)
Occurs when Outlook has completed reading the properties of the item.

## Version information

Version Added: Outlook 2013 


## Syntax

_expression_. `ReadComplete`_(Cancel)_

_expression_ A variable that represents a [TaskRequestAcceptItem](Outlook.TaskRequestAcceptItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the read operation is not completed and the item is not displayed in the Reading Pane or inspector.|

## Remarks

The **ReadComplete** event occurs after the [BeforeRead](Outlook.TaskRequestAcceptItem.BeforeRead.md) event and before the [Read](Outlook.TaskRequestAcceptItem.Read.md) event for the item.

To determine when the item is unloaded from memory, use the [Unload](Outlook.TaskRequestAcceptItem.Unload.md) event.

The **ReadComplete** event corresponds to the Exchange Client Extensions (ECE) event **IExchExtMessageEvents::OnReadComplete**.


## See also


[TaskRequestAcceptItem Object](Outlook.TaskRequestAcceptItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]