---
title: TaskRequestItem.GetAssociatedTask method (Outlook)
keywords: vbaol11.chm1906
f1_keywords:
- vbaol11.chm1906
ms.prod: outlook
api_name:
- Outlook.TaskRequestItem.GetAssociatedTask
ms.assetid: ec170266-9898-79d8-03e9-7ea38d789d40
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestItem.GetAssociatedTask method (Outlook)

Returns a  **[TaskItem](Outlook.TaskItem.md)** object that represents the requested task.


## Syntax

_expression_. `GetAssociatedTask`( `_AddToTaskList_` )

_expression_ A variable that represents a [TaskRequestItem](Outlook.TaskRequestItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _AddToTaskList_|Required| **Boolean**| **True** if the task is added to the default **Tasks** folder.|

## Return value

A  **TaskItem** object that represents the requested task.


## Remarks

The  **GetAssociatedTask** method will not work unless the **TaskItem** is processed before the method is called. To do so, call the **[Display](Outlook.TaskItem.Display.md)** method before calling **GetAssociatedTask**.


## Example

This Microsoft Visual Basic for Applications (VBA) example accepts a  **[TaskRequestItem](Outlook.TaskRequestItem.md)**, sending the response without displaying the inspector.


```vb
Sub AcceptTask() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myTasks As Outlook.Folder 
 
 Dim myNewTaskItem As Outlook.TaskItem 
 
 Dim mytaskreqItem As Outlook.TaskRequestItem 
 
 Dim myItem As Outlook.TaskItem 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 Set myTasks = myNameSpace.GetDefaultFolder(olFolderInbox) 
 
 Set mytaskreqItem = myTasks.Items.Find("[Subject] = ""Meeting w/ Nate Sun""") 
 
 If Not TypeName(mytaskreqItem) = "Nothing" Then 
 
 Set myNewTaskItem = mytaskreqItem.GetAssociatedTask(True) 
 
 Set myItem = myNewTaskItem.Respond(olTaskAccept, True, True) 
 
 myItem.Send 
 
 End If 
 
End Sub
```


## See also


[TaskRequestItem Object](Outlook.TaskRequestItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]