---
title: ContactItem.Read event (Outlook)
ms.prod: outlook
api_name:
- Outlook.ContactItem.Read
ms.assetid: 508b4637-9d74-7645-7719-3c148d0688d8
ms.date: 06/08/2017
localization_priority: Normal
---


# ContactItem.Read event (Outlook)

Occurs when an instance of the parent object is opened for editing by the user. 


## Syntax

_expression_. `Read`

_expression_ A variable that represents a [ContactItem](Outlook.ContactItem.md) object.


## Remarks

The **Read** event differs from the **[Open](Outlook.ContactItem.Open.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an **[Inspector](Outlook.Inspector.md)**.


## See also


[ContactItem Object](Outlook.ContactItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]