---
title: TaskRequestUpdateItem.GetAssociatedTask method (Outlook)
keywords: vbaol11.chm1955
f1_keywords:
- vbaol11.chm1955
ms.prod: outlook
api_name:
- Outlook.TaskRequestUpdateItem.GetAssociatedTask
ms.assetid: b663f5fe-05bf-c1c7-f53b-1fbd308f22f8
ms.date: 06/08/2017
localization_priority: Normal
---


# TaskRequestUpdateItem.GetAssociatedTask method (Outlook)

Returns a  **[TaskItem](Outlook.TaskItem.md)** object that represents the requested task.


## Syntax

_expression_. `GetAssociatedTask`( `_AddToTaskList_` )

_expression_ A variable that represents a [TaskRequestUpdateItem](Outlook.TaskRequestUpdateItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _AddToTaskList_|Required| **Boolean**| **True** if the task is added to the default **Tasks** folder.|

## Return value

A  **TaskItem** object that represents the requested task.


## Remarks

The  **GetAssociatedTask** method will not work unless the **TaskItem** is processed before the method is called. To do so, call the **[Display](Outlook.TaskItem.Display.md)** method before calling **GetAssociatedTask**.


## Example

This Microsoft Visual Basic for Applications (VBA) example accepts a  **[TaskRequestUpdateItem Object](Outlook.TaskRequestUpdateItem.md)**, sending the response without displaying the inspector.


```vb
Sub AcceptTask() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myTasks As Outlook.Folder 
 
 Dim myNewTaskItem As Outlook.TaskItem 
 
 Dim mytaskrequpdateItem As Outlook.TaskRequestUpdateItem 
 
 Dim myItem As Outlook.TaskItem 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 Set myTasks = myNameSpace.GetDefaultFolder(olFolderInbox) 
 
 Set mytaskrequdpateItem = myTasks.Items.Find("[Subject] = ""Meeting w/ Nate Sun""") 
 
 If Not TypeName(mytaskrequpdateItem) = "Nothing" Then 
 
 Set myNewTaskItem = mytaskrequpdateItem.GetAssociatedTask(True) 
 
 Set myItem = myNewTaskItem.Respond(olTaskAccept, True, True) 
 
 myItem.Send 
 
 End If 
 
End Sub
```


## See also


[TaskRequestUpdateItem Object](Outlook.TaskRequestUpdateItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]