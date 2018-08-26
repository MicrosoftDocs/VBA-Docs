---
title: Event Statement
keywords: vblr6.chm1103515
f1_keywords:
- vblr6.chm1103515
ms.prod: office
ms.assetid: 14493dfc-5b73-f870-742a-cd4edcf69899
ms.date: 06/08/2017
---


# Event Statement

Declares a user-defined event.

<<<<<<< HEAD
 **Syntax**
=======
## Syntax
>>>>>>> master

[ **Public** ] **Event**_procedurename_ [ **(**_arglist_**)** ]

The  **Event** statement has these parts:


|**Part**|**Description**|
|:-----|:-----|
<<<<<<< HEAD
|**Public**|Optional. Specifies that the  **Event** visible throughout the[project](../../Glossary/vbe-glossary.md).  **Events** types are **Public** by default. Note that events can only be raised in the[module](../../Glossary/vbe-glossary.md) in which they are declared.|
=======
|**Public**|Optional. Specifies that the  **Event** visible throughout the[project](../../Glossary/vbe-glossary.md#project).  **Events** types are **Public** by default. Note that events can only be raised in the[module](../../Glossary/vbe-glossary.md#module) in which they are declared.|
>>>>>>> master
| _procedurename_|Required. Name of the event; follows standard variable naming conventions.|

The  _arglist_ argument has the following syntax and parts:
[ **ByVal** |**ByRef** ] _varname_ [ **( )** ] [ **As**_type_ ]


|**Part**|**Description**|
|:-----|:-----|
<<<<<<< HEAD
|**ByVal**|Optional. Indicates that the [argument](../../Glossary/vbe-glossary.md) is passed[by value](../../Glossary/vbe-glossary.md).|
|**ByRef**|Optional. Indicates that the argument is passed [by reference](../../Glossary/vbe-glossary.md).  **ByRef** is the default in Visual Basic.|
| _varname_|Required. Name of the variable representing the argument being passed to the [procedure](../../Glossary/vbe-glossary.md); follows standard variable naming conventions.|
| _type_|Optional. [Data type](../../Glossary/vbe-glossary.md) of the argument passed to the procedure; may be[Byte](../../Glossary/vbe-glossary.md), [Boolean](../../Glossary/vbe-glossary.md), [Integer](../../Glossary/vbe-glossary.md), [Long](../../Glossary/vbe-glossary.md), [Currency](../../Glossary/vbe-glossary.md), [Single](../../Glossary/vbe-glossary.md), [Double](../../Glossary/vbe-glossary.md), [Decimal](../../Glossary/vbe-glossary.md) (not currently supported),[Date](../../Glossary/vbe-glossary.md), [String](../../Glossary/vbe-glossary.md) (variable length only),[Object](../../Glossary/vbe-glossary.md), [Variant](../../Glossary/vbe-glossary.md), a [user-defined type](../../Glossary/vbe-glossary.md), or an object type.|

 **Remarks**
Once the event has been declared, use the  **RaiseEvent** statement to fire the event. A syntax error occurs if an **Event** declaration appears in a[standard module](../../Glossary/vbe-glossary.md). An event can't be declared to return a value. A typical event might be declared and raised as shown in the following fragments:
=======
|**ByVal**|Optional. Indicates that the [argument](../../Glossary/vbe-glossary.md#argument) is passed[by value](../../Glossary/vbe-glossary.md#by-value).|
|**ByRef**|Optional. Indicates that the argument is passed [by reference](../../Glossary/vbe-glossary.md#by-reference).  **ByRef** is the default in Visual Basic.|
| _varname_|Required. Name of the variable representing the argument being passed to the [procedure](../../Glossary/vbe-glossary.md#procedure); follows standard variable naming conventions.|
| _type_|Optional. [Data type](../../Glossary/vbe-glossary.md#data-type) of the argument passed to the procedure; may be[Byte](../../Glossary/vbe-glossary.md#byte-data-type), [Boolean](../../Glossary/vbe-glossary.md#boolean-data-type), [Integer](../../Glossary/vbe-glossary.md#integer-data-type), [Long](../../Glossary/vbe-glossary.md#long-data-type), [Currency](../../Glossary/vbe-glossary.md#currency-data-type), [Single](../../Glossary/vbe-glossary.md#single-data-type), [Double](../../Glossary/vbe-glossary.md#double-data-type), [Decimal](../../Glossary/vbe-glossary.md#decimal-data-type) (not currently supported),[Date](../../Glossary/vbe-glossary.md#date-data-type), [String](../../Glossary/vbe-glossary.md#string-data-type) (variable length only),[Object](../../Glossary/vbe-glossary.md#object), [Variant](../../Glossary/vbe-glossary.md#variant-data-type), a [user-defined type](../../Glossary/vbe-glossary.md#user-defined-type), or an object type.|

## Remarks

Once the event has been declared, use the  **RaiseEvent** statement to fire the event. A syntax error occurs if an **Event** declaration appears in a[standard module](../../Glossary/vbe-glossary.md#standard-module). An event can't be declared to return a value. A typical event might be declared and raised as shown in the following fragments:
>>>>>>> master



```vb
' Declare an event at module level of a class module 
 
Event LogonCompleted (UserName as String) 
 
Sub 
 RaiseEvent LogonCompleted("AntoineJan") 
End Sub
```


 **Note**  You can declare event arguments just as you do arguments of procedures, with the following exceptions: events cannot have named arguments,  **Optional** arguments, or **ParamArray** arguments. Events do not have return values.


## Example

The following example uses events to count off seconds during a demonstration of the fastest 100 meter race. The code illustrates all of the event-related methods, properties, and statements, including the  **Event** statement.

The class that raises an event is the event source, and the classes that implement the event are the sinks. An event source can have multiple sinks for the events it generates. When the class raises the event, that event is fired on every class that has elected to sink events for that instance of the object.

The example also uses a form ( `Form1` ) with a button ( `Command1` ), a label ( `Label1` ), and two text boxes ( `Text1` and `Text2` ). When you click the button, the first text box displays "From Now" and the second starts to count seconds. When the full time (9.84 seconds) has elapsed, the first text box displays "Until Now" and the second displays "9.84"

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
 Label1.Caption = "The fastest 100 meter run took this long:" 
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

The remaining code is in a class module named TimerState. The  **Event** statements declare the procedures initiated when events are raised.




```vb
Option Explicit 
Public Event UpdateTime(ByVal dblJump As Double)Public Event ChangeText() 
 
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


