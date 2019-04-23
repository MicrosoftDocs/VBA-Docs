---
title: ReportItem.Read event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ReportItem.Read
ms.assetid: 7b142bcb-dd96-a0ec-5684-b7311f34d772
ms.date: 06/08/2017
localization_priority: Normal
---


# ReportItem.Read event (Outlook)

Occurs when an instance of the parent object is opened for editing by the user. 


## Syntax

_expression_. `Read`

_expression_ A variable that represents a [ReportItem](Outlook.ReportItem.md) object.


## Remarks

The  **Read** event differs from the **[Open](Outlook.ReportItem.Open.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an **[Inspector](Outlook.Inspector.md)**.


## See also


[ReportItem Object](Outlook.ReportItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]