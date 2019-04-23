---
title: Form.Module property (Access)
keywords: vbaac10.chm13500
f1_keywords:
- vbaac10.chm13500
ms.prod: access
api_name:
- Access.Form.Module
ms.assetid: f4583bc6-a412-811e-a428-cfa10a911d35
ms.date: 03/13/2019
localization_priority: Normal
---


# Form.Module property (Access)

You can use the **Module** property to specify a form module. Read-only **Module** object.


## Syntax

_expression_.**Module**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

The **Module** property also returns a reference to a specified **Module** object. Use the **Module** property to access the properties and methods of a **Module** object associated with a **Form** or **Report** object.

The setting of the **HasModule** property of a form or report determines whether it has an associated module. 

If the **HasModule** property is **False**, the form or report does not have an associated module. When you refer to the **Module** property of that form or report while in Design view, Microsoft Access creates the associated module and sets the **HasModule** property to **True**. 

If you refer to the **Module** property of a form or report at run-time and the object has its **HasModule** property set to **False**, an error will occur.

You could use this property with any of the properties and methods of the module object.


## Example

The following example uses the **Module** property to insert the **Beep** method in a form's **Open** event.

```vb
Dim strFormOpenCode As String 
Dim mdl As Module 
 
Set mdl = Forms!MyForm.Module 
strFormOpenCode = "Sub Form_Open(Cancel As Integer)" _ 
 & vbCrLf & "Beep" & vbCrLf & "End Sub" 
 With mdl 
 .InsertText strFormOpenCode 
 End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]