---
title: Masters.CellChanged event (Visio)
keywords: vis_sdr.chm10819090
f1_keywords:
- vis_sdr.chm10819090
ms.prod: visio
api_name:
- Visio.Masters.CellChanged
ms.assetid: 0abb97fc-ffd6-02ef-b9b3-bbad421c1daf
ms.date: 06/08/2017
localization_priority: Normal
---


# Masters.CellChanged event (Visio)

Occurs after the value changes in a cell in a document.


## Syntax

_expression_.**CellChanged** (_Cell_)

_expression_ A variable that represents a **[Masters](Visio.Masters.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cell_|Required| **[IVCELL]**|The cell whose value has changed.|

## Remarks

If you are using Microsoft Visual Basic or Visual Basic for Applications (VBA), the syntax in this topic describes a common, efficient way to handle events.

If you want to create your own **Event** objects, use the **[Add](visio.eventlist.add.md)** or **[AddAdvise](visio.eventlist.addadvise.md)** method. 

To create an **Event** object that runs an add-on, use the **Add** method as it applies to the **EventList** collection. 

To create an **Event** object that receives notification, use the **AddAdvise** method. 

To find an event code for the event that you want to create, see [Event codes](../visio/Concepts/event-codesvisio.md).




> [!NOTE] 
> You can use VBA  **WithEvents** variables to sink the **CellChanged** event.

For performance considerations, the  **Document** object's event set does not include the **CellChanged** event. To sink the **CellChanged** event from a **Document** object (and the **[ThisDocument](../visio/Concepts/about-the-thisdocument-object-visio.md)** object in a VBA project), you must use the **AddAdvise** method.


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
 
 Debug.Print vsoCell.Shape.Name & " " & vsoCell.Name & " changed to =" & vsoCell.Formula 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]