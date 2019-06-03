---
title: CommandBarButton.Tag property (Office)
ms.prod: office
api_name:
- Office.CommandBarButton.Tag
ms.assetid: c73a12a8-8b20-1e32-ad98-ae0bb3b1daed
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarButton.Tag property (Office)

Gets or sets information about the **CommandBarButton** control, such as data that can be used as an argument in procedures, or information that identifies the control. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Tag**

_expression_ A variable that represents a **[CommandBarButton](Office.CommandBarButton.md)** object.


## Return value

String


## Example

To avoid duplicate calls of the same class when triggered with events, define the **Tag** property unique to the events. The following example demonstrates this concept with two modules.


```vb
Public WithEvents oBtn As CommandBarButton 
 
Private Sub oBtn_click(ByVal ctrl As Office.CommandBarButton, CancelDefault As Boolean) 
    MsgBox "Clicked " & ctrl.Caption 
 
End Sub 
 
Dim oBtns As New Collection 
      
Sub Use_Tag() 
     
    Dim oEvt As CBtnEvent 
    Set oBtns = Nothing 
 
    For i = 1 To 5 
        Set oEvt = New CBtnEvent 
        Set oEvt.oBtn = Application.CommandBars("Worksheet Menu Bar").Controls.Add(msoControlButton) 
        With oEvt.oBtn 
            .Caption = "Btn" & i 
            .Style = msoButtonCaption 
            .Tag = "Hello" & i 
        End With 
        oBtns.Add oEvt 
    Next 
      
End Sub
```

<br/>

This example sets the tag for the button on the custom command bar to **Spelling Button** and displays the tag in a message box.

```vb
CommandBars("Custom").Controls(1).Tag = "Spelling Button" 
MsgBox (CommandBars("Custom").Controls(1).Tag)
```


## See also

- [CommandBarButton object members](overview/library-reference/commandbarbutton-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]