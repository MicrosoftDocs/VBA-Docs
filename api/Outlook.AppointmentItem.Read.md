---
title: AppointmentItem.Read event (Outlook)
ms.prod: outlook
api_name:
- Outlook.AppointmentItem.Read
ms.assetid: aa39ec06-19ed-4655-6990-e4c4c45649d5
ms.date: 06/08/2017
localization_priority: Normal
---


# AppointmentItem.Read event (Outlook)

Occurs when an instance of the parent object is opened for editing by the user. 


## Syntax

_expression_. `Read`

_expression_ A variable that represents an [AppointmentItem](Outlook.AppointmentItem.md) object.


## Remarks

The **Read** event differs from the **[Open](Outlook.AppointmentItem.Open.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an **[Inspector](Outlook.Inspector.md)**.


## See also


[AppointmentItem Object](Outlook.AppointmentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]