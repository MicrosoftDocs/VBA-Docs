---
title: PostItem.Read event (Outlook)
ms.prod: outlook
api_name:
- Outlook.PostItem.Read
ms.assetid: 404c9b17-c5b6-a802-420a-f8fd279b5f9b
ms.date: 06/08/2017
localization_priority: Normal
---


# PostItem.Read event (Outlook)

Occurs when an instance of the parent object is opened for editing by the user. 


## Syntax

_expression_. `Read`

_expression_ A variable that represents a [PostItem](Outlook.PostItem.md) object.


## Remarks

The **Read** event differs from the **[Open](Outlook.PostItem.Open.md)** event in that **Read** occurs whenever the user selects the item in a view that supports in-cell editing as well as when the item is being opened in an **[Inspector](Outlook.Inspector.md)**.


## See also


[PostItem Object](Outlook.PostItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]