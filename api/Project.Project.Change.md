---
title: Project.Change event (Project)
ms.prod: project-server
api_name:
- Project.Project.Change
ms.assetid: ef109b59-c7be-0707-9716-13c86180c27c
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.Change event (Project)

Occurs when a change is made to data in the project. An action affecting several items at once is considered to be one change.


## Syntax

_expression_. `Change`( `_pj_` )

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**|The project that changed.|

## Return value

**Nothing**


## Remarks

The **Change** event does not occur for actions such as switching views, applying filters, changing formatting, and so on.

Project events do not occur when the project is embedded in another document or application. 


## Example

 The following example shows how the **ProjectTaskNew** event can trap project-level events. In this case, the **App_ProjectTaskNew** event handler sets the global **ProjTaskNew** variable that the **Change** event handler uses. You can use similar code with the **[ProjectResourceNew](Project.Application.ProjectResourceNew.md)** and **[ProjectAssignmentNew](Project.Application.ProjectAssignmentNew.md)** events.


1. Create a new class module named  **EventClassModule**, and then insert the following code:
    
  ```vb
  Option Explicit 
Option Base 1 
 
Public WithEvents App As Application 
Public WithEvents Proj As Project 
 
Dim NewTaskIDs() As Integer 
Dim NumNewTasks As Integer 
 
Dim ProjTaskNew As Boolean 
 
Private Sub App_ProjectTaskNew(ByVal pj As Project, ByVal ID As Long) 
    NumNewTasks = NumNewTasks + 1 
 
    If ProjTaskNew Then 
        ReDim Preserve NewTaskIDs(NumNewTasks) As Integer 
    Else 
        ReDim NewTaskIDs(NumNewTasks) As Integer 
    End If 
 
    NewTaskIDs(NumNewTasks) = ID 
 
    ProjTaskNew = True 
End Sub 
 
Private Sub Proj_Change(ByVal pj As Project) 
 Dim NewTaskID As Variant 
 
    If ProjTaskNew Then 
        For Each NewTaskID In NewTaskIDs 
            MsgBox "New Task Name: " & ActiveProject.Tasks.UniqueID(NewTaskID).Name 
        Next NewTaskID 
 
        NumNewTasks = 0 
 
        ProjTaskNew = False 
    End If 
End Sub
```


    
    
2. In a separate module, insert the following code:
    
  ```vb
  Option Explicit 
 
Dim X As New EventClassModule 
 
Sub Initialize_App() 
    Set X.App = MSProject.Application 
    Set X.Proj = Application.ActiveProject 
End Sub
  ```


    
    
3. Run the **Initialize_App** procedure to start listening to the events.
    
4. Create a new task. The event handler shows a message box every time a new task is added.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]