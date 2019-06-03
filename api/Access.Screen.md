---
title: Screen object (Access)
keywords: vbaac10.chm12484
f1_keywords:
- vbaac10.chm12484
ms.prod: access
api_name:
- Access.Screen
ms.assetid: 00743775-071b-9ccd-7687-f3b992e9346e
ms.date: 03/21/2019
localization_priority: Normal
---


# Screen object (Access)

The **Screen** object refers to the particular form, report, or control that currently has the focus.


## Remarks

You can use the **Screen** object together with its properties to refer to a particular form, report, or control that has the focus.

For example, you can use the **Screen** object with the **ActiveForm** property to refer to the form in the active window without knowing the form's name. The following example displays the name of the form in the active window.

```vb
MsgBox Screen.ActiveForm.Name
```

Referring to the **Screen** object doesn't make a form, report, or control active. To make a form, report, or control active, you must use the **SelectObject** method of the **[DoCmd](Access.DoCmd.md)** object.

If you refer to the **Screen** object when there's no active form, report, or control, Microsoft Access returns a run-time error. For example, if a standard module is in the active window, the code in the preceding example would return an error.


## Example

The following example uses the **Screen** object to print the name of the form in the active window and of the active control on that form.

```vb
Sub ActiveObjects() 
 Dim frm As Form, ctl As Control 
 
 ' Return Form object pointing to active form. 
 Set frm = Screen.ActiveForm 
 MsgBox frm.Name & " is the active form." 
 ' Return Control object pointing to active control. 
 Set ctl = Screen.ActiveControl 
 MsgBox ctl.Name & " is the active control " _ 
 & "on this form." 
End Sub 

```


## Properties

- [ActiveControl](Access.Screen.ActiveControl.md)
- [ActiveDatasheet](Access.Screen.ActiveDatasheet.md)
- [ActiveForm](Access.Screen.ActiveForm.md)
- [ActiveReport](Access.Screen.ActiveReport.md)
- [Application](Access.Screen.Application.md)
- [MousePointer](Access.Screen.MousePointer.md)
- [Parent](Access.Screen.Parent.md)
- [PreviousControl](Access.Screen.PreviousControl.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
