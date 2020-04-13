---
title: DocumentItem.Read event (Outlook)
ms.prod: outlook
api_name:
- Outlook.DocumentItem.Read
ms.assetid: da5e82e6-43b9-d040-e529-2388049a8e1b
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentItem.Read event (Outlook)

Occurs when an instance of the parent object is opened for editing by the user. 


## Syntax

_expression_. `Read`

_expression_ A variable that represents a [DocumentItem](Outlook.DocumentItem.md) object.


## Remarks

The **Read** event differs from the **[Open](Outlook.DocumentItem.Open.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an **[Inspector](Outlook.Inspector.md)**.


## See also


[DocumentItem Object](Outlook.DocumentItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]