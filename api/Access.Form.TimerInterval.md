---
title: Form.TimerInterval property (Access)
keywords: vbaac10.chm13462
f1_keywords:
- vbaac10.chm13462
ms.prod: access
api_name:
- Access.Form.TimerInterval
ms.assetid: ee56bcf8-20cb-9d86-ed17-3b85ac88f6f1
ms.date: 03/15/2019
localization_priority: Normal
---


# Form.TimerInterval property (Access)

You can use the **TimerInterval** property to specify the interval, in milliseconds, between **[Timer](Access.Form.Timer.md)** events on a form. Read/write **Long**.


## Syntax

_expression_.**TimerInterval**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

The **TimerInterval** property setting is a **Long Integer** value between 0 and 2,147,483,647.

You can set this property by using the form's property sheet, a macro, or Visual Basic.

> [!NOTE] 
> When using Visual Basic, you set the **TimerInterval** property in the form's **Load** event.

To run Visual Basic code at intervals specified by the **TimerInterval** property, put the code in the form's **Timer** event procedure. For example, to requery records every 30 seconds, put the code to requery the records in the form's **Timer** event procedure, and then set the **TimerInterval** property to 30000.

## Example

The following example shows how to create a flashing button on a form by displaying and hiding an icon on the button. The form's **Load** event procedure sets the form's **TimerInterval** property to 1000 so that the icon display is toggled once every second.

```vb
Sub Form_Load() 
    Me.TimerInterval = 1000 
End Sub 
 
Sub Form_Timer() 
    Static intShowPicture As Integer 
    If intShowPicture Then 
        ' Show icon. 
        Me!btnPicture.Picture = "C:\Icons\Flash.ico" 
    Else 
        ' Don't show icon. 
        Me!btnPicture.Picture = "" 
    End If 
    intShowPicture = Not intShowPicture 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
