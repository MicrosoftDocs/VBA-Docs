---
title: Application.IsInScope Property (Visio)
keywords: vis_sdr.chm10013750
f1_keywords:
- vis_sdr.chm10013750
ms.prod: visio
api_name:
- Visio.Application.IsInScope
ms.assetid: adb9a52f-8e62-9d92-d8bf-81bed48b2cc3
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.IsInScope Property (Visio)

Determines whether a call to an event handler is between an  **EnterScope** event and an **ExitScope** event for a scope. Read-only.


## Syntax

 _expression_. `IsInScope`( `_nCmdID_` )

 _expression_ A variable that represents an [Application](./Visio.Application.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _nCmdID_|Required| **Long**|The scope ID.|

## Return value

Boolean


## Remarks

Constants representing scope IDs are prefixed with  **visCmd** and are declared by the Visio type library. You can also use an ID returned by the **BeginUndoScope** method.

You could use this property in a  **CellChanged** event handler to determine whether a cell change was the result of a particular operation.


## Example

This example shows how to use the  **IsInScope** property to determine whether a call to a procedure that handles the **CellChanged** event is in a particular scope?that is, whether the call occurs between the **EnterScope** and **ExitScope** events for that scope.


```vb
 
Private WithEvents vsoApplication As Visio.Application 
Private lngScopeID As Long 
 
Public Sub IsInScope_Example() 
 
 Dim vsoShape As Visio.Shape 
 
 'Set the module-level application variable to 
 'trap application-level events. 
 Set vsoApplication = Application 
 
 'Begin a scope. 
 lngScopeID = Application.BeginUndoScope("Draw Shapes") 
 
 'Draw three shapes. 
 Set vsoShape = ActivePage.DrawRectangle(1, 2, 2, 1) 
 ActivePage.DrawOval 3, 4, 4, 3 
 ActivePage.DrawLine 4, 5, 5, 4 
 
 'Change a cell (to trigger the CellChanged event). 
 vsoShape.Cells("Width").Formula = 5 
 
 'End and commit this scope. 
 Application.EndUndoScope lngScopeID, True 
 
End Sub 
 
Private Sub vsoApplication_CellChanged(ByVal Cell As IVCell) 
 
 'Check to see if this cell change is the result of something 
 'happening within the scope. 
 If vsoApplication.IsInScope(lngScopeID) Then 
 Debug.Print Cell.Name & " changed in scope "; lngScopeID 
 End If 
 
End Sub 
 
Private Sub vsoApplication_EnterScope(ByVal app As IVApplication, _ 
 ByVal nScopeID As Long, _ 
 ByVal bstrDescription As String) 
 
 If vsoApplication.CurrentScope = lngScopeID Then 
 Debug.Print "Entering my scope " & nScopeID 
 Else 
 Debug.Print "Enter Scope " & bstrDescription & "(" & nScopeID & ")" 
 End If 
 
End Sub 
 
Private Sub vsoApplication_ExitScope(ByVal app As IVApplication, _ 
 ByVal nScopeID As Long, _ 
 ByVal bstrDescription As String, _ 
 ByVal bErrOrCancelled As Boolean) 
 
 If vsoApplication.CurrentScope = lngScopeID Then 
 Debug.Print "Exiting my scope " & nScopeID 
 Else 
 Debug.Print "ExitScope " & bstrDescription & "(" & nScopeID & ")" 
 End If 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]