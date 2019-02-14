---
title: Screen.ActiveForm property (Access)
keywords: vbaac10.chm12490
f1_keywords:
- vbaac10.chm12490
ms.prod: access
api_name:
- Access.Screen.ActiveForm
ms.assetid: 5cf41661-656e-e62f-530e-0d2fa5466146
ms.date: 06/08/2017
localization_priority: Priority
---


# Screen.ActiveForm property (Access)

You can use the  **ActiveForm** property together with the **Screen** object to identify or refer to the form that has the focus. Read-only **Form** object.


## Syntax

_expression_. `ActiveForm`

_expression_ A variable that represents a **[Screen](Access.Screen.md)** object.


## Remarks

This property setting contains a reference to the  **[Form](Access.Form.md)** object that has the focus at run time.

You can use the  **ActiveForm** property to refer to an active form together with one of its properties or methods. The following example displays the **Name** property setting of the active form.




```vb
Dim frmCurrentForm As Form 
Set frmCurrentForm = Screen.ActiveForm 
MsgBox "Current form is " & frmCurrentForm.Name
```

If a subform has the focus,  **ActiveForm** refers to the main form. If no form or subform has the focus when you use the **ActiveForm** property, an error occurs.


## See also


[Screen Object](Access.Screen.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]