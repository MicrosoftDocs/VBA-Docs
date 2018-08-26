---
title: InvisibleApp.CellChanged Event (Visio)
ms.prod: visio
api_name:
- Visio.InvisibleApp.CellChanged
ms.assetid: 6fbb3f63-db6f-6937-7ace-97bf9241615a
ms.date: 06/08/2017
---


# InvisibleApp.CellChanged Event (Visio)

Occurs after the value changes in a cell in a document.


## Syntax

Private Sub  _expression_ _'CellChanged'(**_ByVal Cell As [IVCELL]_**)

 _expression_ A variable that represents an [InvisibleApp](./Visio.InvisibleApp.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cell_|Required| **[IVCELL]**|The cell whose value has changed.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own  **Event** objects, use the **Add** or **AddAdvise** method. To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. To create an **Event** object that receives notification, use the **AddAdvise** method. To find an event code for the event you want to create, see[Event codes](../visio/Concepts/event-codesvisio.md).




 **Note**  You can use VBA  **WithEvents** variables to sink the **CellChanged** event.

For performance considerations, the  **Document** object's event set does not include the **CellChanged** event. To sink the **CellChanged** event from a **Document** object (and the **ThisDocument** object in a VBA project), you must use the **AddAdvise** method.


## Example

This VBA module shows how to use the  **CellChanged** event to trap changes to a shape's cells.


```vb
 
Private WithEvents vsoApplication As Visio.Application 
 
Public Sub CellChanged_Example() 
 
 Dim vsoShape As Visio.Shape 
 
 'Set a module-level variable to trap application-level events. 
 Set vsoApplication = Application 
 
 'Draw a shape. 
 Set vsoShape = ActivePage.DrawRectangle(1, 2, 2, 1) 
 
 'Change a cell (to trigger a CellChanged event). 
 vsoShape.Cells("Width").Formula = 5 
 
End Sub 
 
Private Sub vsoApplication_CellChanged(ByVal vsoCell As IVCell) 
 
<<<<<<< HEAD
 Debug.Print vsoCell.Shape.Name &; " " &; vsoCell.Name &; " changed to =" &; vsoCell.Formula 
=======
 Debug.Print vsoCell.Shape.Name & " " & vsoCell.Name & " changed to =" & vsoCell.Formula 
>>>>>>> master
 
End Sub
```


