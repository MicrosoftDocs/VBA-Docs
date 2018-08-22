---
title: RaiseEvent Statement
keywords: vblr6.chm1103516
f1_keywords:
- vblr6.chm1103516
ms.prod: office
ms.assetid: 4de2ad26-cb93-19b1-9f44-e6c1b5d619f3
ms.date: 06/08/2017
---


# RaiseEvent Statement

Fires an event declared at [module level](../../Glossary/vbe-glossary.md#module-level) within a[class](../../Glossary/vbe-glossary.md#class), form, or document.

## Syntax

**RaiseEvent**_eventname_ [ **(**_argumentlis_ t **)** ]

The required  _eventname_ is the name of an event declared within the[module](../../Glossary/vbe-glossary.md#module) and follows Basic variable naming conventions.
The  **RaiseEvent** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _eventname_|Required. Name of the event to fire.|
| _argumentlist_|Optional. Comma-delimited list of [variables](../../Glossary/vbe-glossary.md#variable), [arrays](../../Glossary/vbe-glossary.md#array), or [expressions](../../Glossary/vbe-glossary.md#expression) The _argumentlist_ must be enclosed by parentheses. If there are no[arguments](../../Glossary/vbe-glossary.md#argument), the parentheses must be omitted.|

## Remarks

<<<<<<< HEAD
=======
## Remarks

>>>>>>> 54e0a75f224118db0d26fc9363ad519ad35ec788
If the event has not been declared within the module in which it is raised, an error occurs. The following fragment illustrates an event declaration and a procedure in which the event is raised.



```vb
' Declare an event at module level of a class module 
Event LogonCompleted (UserName as String) 
 
Sub 
 ' Raise the event. 
 RaiseEvent LogonCompleted ("AntoineJan") 
End Sub
```

If the event has no arguments, including empty parentheses, in the  **RaiseEvent**, invocation of the event causes an error. You can't use **RaiseEvent** to fire events that are not explicitly declared in the module. For example, if a form has a Click event, you can't fire its Click event using **RaiseEvent**. If you declare a Click event in the[form module](../../Glossary/vbe-glossary.md#form-module), it shadows the form's own Click event. You can still invoke the form's Click event using normal syntax for calling the event, but not using the  **RaiseEvent** statement.
Event firing is done in the order that the connections are established. Since events can have  **ByRef** parameters, a process that connects late may receive parameters that have been changed by an earlier event handler.

## Example

The following example uses events to count off seconds during a demonstration of the fastest 100 meter race. The code illustrates all of the event-related methods, properties, and statements, including the  **RaiseEvent** statement.

The class that raises an event is the event source, and the classes that implement the event are the sinks. An event source can have multiple sinks for the events it generates. When the class raises the event, that event is fired on every class that has elected to sink events for that instance of the object.

The example also uses a form ( `Form1`) with a button ( `Command1`), a label ( `Label1`), and two text boxes ( `Text1` and `Text2`). When you click the button, the first text box displays "From Now" and the second starts to count seconds. When the full time (9.84 seconds) has elapsed, the first text box displays "Until Now" and the second displays "9.84"

The code for specifies the initial and terminal states of the form. It also contains the code executed when events are raised.




```vb
Option Explicit 
 
Private WithEvents mText As TimerState 
 
Private Sub Command1_Click() 
 Text1.Text = "From Now" 
 Text1.Refresh 
 Text2.Text = "0" 
 Text2.Refresh 
 Call mText.TimerTask(9.84) 
End Sub 
 
Private Sub Form_Load() 
 Command1.Caption = "Click to Start Timer" 
 Text1.Text = "" 
 Text2.Text = "" 
 Label1.Caption = "The fastest 100 meters ever run took this long:" 
 Set mText = New TimerState 
 End Sub 
 
Private Sub mText_ChangeText() 
 Text1.Text = "Until Now" 
 Text2.Text = "9.84" 
End Sub 
 
Private Sub mText_UpdateTime(ByVal dblJump As Double) 
 Text2.Text = Str(Format(dblJump, "0")) 
 DoEvents 
End Sub
```

The remaining code is in a class module named TimerState. Included among the commands in this module are the  **Raise Event** statements.




```vb
Option Explicit 
Public Event UpdateTime(ByVal dblJump As Double) 
Public Event ChangeText() 
 
Public Sub TimerTask(ByVal Duration As Double) 
 Dim dblStart As Double 
 Dim dblSecond As Double 
 Dim dblSoFar As Double 
 dblStart = Timer 
 dblSoFar = dblStart 
 
 Do While Timer < dblStart + Duration 
 If Timer - dblSoFar >= 1 Then 
 dblSoFar = dblSoFar + 1 
 RaiseEvent UpdateTime(Timer - dblStart) 
 End If 
 Loop 
 
 RaiseEvent ChangeText 
 
End Sub
```


