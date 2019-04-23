---
title: TabControl.BorderColor property (Access)
keywords: vbaac10.chm10839
f1_keywords:
- vbaac10.chm10839
ms.prod: access
api_name:
- Access.TabControl.BorderColor
ms.assetid: 26a4f80e-d4d2-a633-0e7d-67a2d1143af6
ms.date: 02/20/2019
localization_priority: Normal
---


# TabControl.BorderColor property (Access)

You can use the **BorderColor** property to specify the color of a control's border. Read/write **Long**.


## Syntax

_expression_.**BorderColor**

_expression_ A variable that represents a **[TabControl](Access.TabControl.md)** object.


## Remarks

The **BorderColor** property setting is a numeric expression that corresponds to the color that you want to use for a control's border.

You can set the default for this property by using a control's default control style or the **[DefaultControl](access.form.defaultcontrol.md)** property in Visual Basic.

A control's border color is visible only when its **SpecialEffect** property is set to Flat or Shadowed. If the **SpecialEffect** property is set to something other than Flat or Shadowed, setting the **BorderColor** property changes the **SpecialEffect** property setting to Flat.


## Example

The following example uses the **RGB** function to set the **BorderColor**, **BackColor**, and **ForeColor** properties depending on the value of the **txtPastDue** text box. You can also use the **QBColor** function to set these properties. 

Putting the following code in the **Form_Current( )** event sets the control display characteristics as soon as the user opens a form or moves to a new record.


```vb
Sub Form_Current() 
 Dim curAmntDue As Currency, lngBlack As Long 
 Dim lngRed As Long, lngYellow As Long, lngWhite As Long 
 
 If Not IsNull(Me!txtPastDue.Value) Then 
 curAmntDue = Me!txtPastDue.Value 
 Else 
 Exit Sub 
 End If 
 lngRed = RGB(255, 0, 0) 
 lngBlack = RGB(0, 0, 0) 
 lngYellow = RGB(255, 255, 0) 
 lngWhite = RGB(255, 255, 255) 
 If curAmntDue > 100 Then 
 Me!txtPastDue.BorderColor = lngRed 
 Me!txtPastDue.ForeColor = lngRed 
 Me!txtPastDue.BackColor = lngYellow 
 Else 
 Me!txtPastDue.BorderColor = lngBlack 
 Me!txtPastDue.ForeColor = lngBlack 
 Me!txtPastDue.BackColor = lngWhite 
 End If 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]