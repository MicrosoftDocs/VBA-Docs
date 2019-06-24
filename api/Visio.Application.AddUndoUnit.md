---
title: Application.AddUndoUnit method (Visio)
keywords: vis_sdr.chm10016075
f1_keywords:
- vis_sdr.chm10016075
ms.prod: visio
api_name:
- Visio.Application.AddUndoUnit
ms.assetid: 90542078-5efa-fec6-b853-41f8a998bea9
ms.date: 06/24/2019
localization_priority: Normal
---


# Application.AddUndoUnit method (Visio)

Adds an object that supports the **IOleUndoUnit** or **IVBUndoUnit** interface to the Microsoft Visio undo queue.


## Syntax

_expression_.**AddUndoUnit** (_pUndoUnit_)

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pUndoUnit_|Required| **[UNKNOWN]**|A reference to an object that supports the **IOleUndoUnit** or **IVBUndoUnit** interface.|

## Return value

Nothing


## Remarks

For information about implementing the **IOleUndoUnit** interface on your object, see the [Microsoft Platform SDK](https://developer.microsoft.com/windows/downloads/windows-10-sdk). For information about implementing the **IVBUndoUnit** interface, see [Developing Microsoft Visio Solutions](https://developer.microsoft.com/visio).


## Example

The following procedure shows how to use the **AddUndoUnit** method to add an object to the Visio undo queue. When a shape is added to the active document, the procedure checks to see if it was added as a result of an undo or redo action, and if not, it adds an Undo unit.

This procedure is a member of class **clsParticipateInUndo**, which is defined in one of two related class modules in the Code Samples Library in the Visio SDK, and is not intended to be run independently. (The other class module defines class **clsVBUndoUnits**.) For more information about these class modules, see the [Visio SDK](https://www.microsoft.com/download/details.aspx?id=51221).

```vb
 
Private Sub mvsoDocument_ShapeAdded(ByVal vsoShape As IVShape) 
 
 Dim VBUndoUnit As clsVBUndoUnits 
 
 On Error GoTo mvsoDocument_ShapeAdded_Err 
 
 If Not (mvsoApplication Is Nothing) Then 
 
 If Not msvoApplication.IsUndoingOrRedoing Then 
 
 'Increment the count of undoable actions. 
 IncrementModuleVar 
 Debug.Print "Original Do: GetModuleVar = " & GetModuleVar 
 
 'Instantiate clsVBUndoUnit, a 
 'class that implements Visio.IVBUndoUnit. 
 Set VBUndoUnit = New clsVBUndoUnits 
 
 'Pass the current instance of the class 
 'of which this procedure is a member, 
 'clsParticipateInUndo, to the Undo unit. 
 VBUndoUnit.SetModelObject Me 
 
 'Add an Undo unit. 
 mvsoApplication.AddUndoUnit VBUndoUnit 
 
 End If 
 
 End If 
 
Exit Sub 
 
mvsoDocument_ShapeAdded_Err: 
 
 MsgBox Err.Description 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]