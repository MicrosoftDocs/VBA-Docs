---
title: JournalItem.Open event (Outlook)
ms.prod: outlook
api_name:
- Outlook.JournalItem.Open
ms.assetid: d564c47d-00d8-29cf-244c-4a2ef1568e26
ms.date: 06/08/2017
localization_priority: Normal
---


# JournalItem.Open event (Outlook)

Occurs when an instance of the parent object is being opened in an **[Inspector](Outlook.Inspector.md)**.


## Syntax

_expression_.**Open** (_Cancel_)

_expression_ A variable that represents a [JournalItem](Outlook.JournalItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required| **Boolean**|(Not used in VBScript).  **False** when the event occurs. If the event procedure sets this argument to **True**, the open operation is not completed and the inspector is not displayed.|

## Remarks

When this event occurs, the  **Inspector** object is initialized but not yet displayed. The **Open** event differs from the **[Read](Outlook.AppointmentItem.Read.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an inspector.

In Microsoft Visual Basic Scripting Edition (VBScript), if you set the return value of this function to  **False**, the open operation is not completed and the inspector is not displayed.


## See also


[JournalItem Object](Outlook.JournalItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]