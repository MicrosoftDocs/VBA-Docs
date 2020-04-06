---
title: Recipient.Resolve method (Outlook)
keywords: vbaol11.chm2358
f1_keywords:
- vbaol11.chm2358
ms.prod: outlook
api_name:
- Outlook.Recipient.Resolve
ms.assetid: 2c4f9243-2e31-642e-78a7-fe74cd73b385
ms.date: 06/08/2017
localization_priority: Normal
---


# Recipient.Resolve method (Outlook)

Attempts to resolve a **[Recipient](Outlook.Recipient.md)** object against the Address Book.


## Syntax

_expression_. `Resolve`

_expression_ A variable that represents a [Recipient](Outlook.Recipient.md) object.


## Return value

 **True** if the object was resolved; otherwise, **False**.


## Example

This Visual Basic for Applications (VBA) example uses  **[CreateItem](Outlook.Application.CreateItem.md)** to create a simple task and delegate it as a task request to another user. Before running this example, replace 'Dan Wilson' with a valid recipient name.


```vb
Sub AssignTask() 
 
 Dim myItem As Outlook.TaskItem 
 
 Dim myDelegate As Outlook.Recipient 
 
 
 
 Set MyItem = Application.CreateItem(olTaskItem) 
 
 MyItem.Assign 
 
 Set myDelegate = MyItem.Recipients.Add("Dan Wilson") 
 
 myDelegate.Resolve 
 
 If myDelegate.Resolved Then 
 
 myItem.Subject = "Prepare Agenda For Meeting" 
 
 myItem.DueDate = Now + 30 
 
 myItem.Display 
 
 myItem.Send 
 
 End If 
 
End Sub
```


## See also


[Recipient Object](Outlook.Recipient.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
