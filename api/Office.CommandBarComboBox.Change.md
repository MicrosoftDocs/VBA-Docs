---
title: CommandBarComboBox.Change event (Office)
keywords: vbaof11.chm229001
f1_keywords:
- vbaof11.chm229001
ms.prod: office
api_name:
- Office.CommandBarComboBox.Change
ms.assetid: ddf1a306-c299-36d5-9851-04d6e5185db9
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarComboBox.Change event (Office)

Occurs when the end user changes the selection in a **CommandBar** combo box.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Change** (_ByVal Ctrl As CommandBarComboBox_)

_expression_ A variable that represents a **[CommandBarComboBox](Office.CommandBarComboBox.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Ctrl_|Required|**CommandBarComboBox**|Represents a **CommandBar** combo box.|

## Remarks

The **Change** event is recognized by the **CommandBarComboBox** object. To return the **Change** event for a particular **CommandBarComboBox** control, use the **WithEvents** keyword to declare a variable, and then set the variable to the **CommandBarComboBox** control. When the **Change** event is triggered, it executes the macro or code that you specified with the **OnAction** property of the control.


## Example

The following example creates a command bar with a **CommandBarComboBox** control containing four selections. The combo box responds to user interaction through the **CommandBarComboBox_Change** event.


```vb
Private ctlComboBoxHandler As New ComboBoxHandler 
Sub AddComboBox() 
 
    Set HostApp = Application 
             
    Dim newBar As Office.CommandBar 
    Set newBar = HostApp.CommandBars.Add(Name:="Test CommandBar", Temporary:=True) 
    Dim newCombo As Office.CommandBarComboBox 
    Set newCombo = newBar.Controls.Add(msoControlComboBox) 
    With newCombo 
        .AddItem "First Class", 1 
        .AddItem "Business Class", 2 
        .AddItem "Coach Class", 3 
        .AddItem "Standby", 4 
        .DropDownLines = 5 
        .DropDownWidth = 75 
        .ListHeaderCount = 0 
    End With 
    ctlComboBoxHandler.SyncBox newCombo 
    newBar.Visible = True  
     
 
End Sub
```

<br/>

The preceding example relies on the following code, which is stored in a class module in the VBA project.

```vb
Private WithEvents ComboBoxEvent As Office.CommandBarComboBox 
Public Sub SyncBox(box As Office.CommandBarComboBox) 
    Set ComboBoxEvent = box 
    If Not box Is Nothing Then 
        MsgBox "Synced " &amp; box.Caption &amp; " ComboBox events." 
    End If 
     
End Sub 
 
Private Sub Class_Terminate() 
    Set ComboBoxEvent = Nothing 
End Sub 
 
Private Sub ComboBoxEvent_Change(ByVal Ctrl As Office.CommandBarComboBox) 
    Dim stComboText As String 
     
    stComboText = Ctrl.Text 
     
        Select Case stComboText 
        Case "First Class" 
            FirstClass 
        Case "Business Class" 
            BusinessClass 
        Case "Coach Class" 
            CoachClass 
        Case "Standby" 
            Standby 
    End Select 
 
End Sub 
Private Sub FirstClass() 
    MsgBox "You selected First Class reservations" 
End Sub 
Private Sub BusinessClass() 
    MsgBox "You selected Business Class reservations" 
End Sub 
Private Sub CoachClass() 
    MsgBox "You selected Coach Class reservations" 
End Sub 
Private Sub Standby() 
    MsgBox "You chose to fly standby" 
End Sub
```


## See also

- [CommandBarComboBox object members](overview/library-reference/commandbarcombobox-members-office.md)
