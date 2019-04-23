---
title: CommandButton.Hyperlink property (Access)
keywords: vbaac10.chm10446
f1_keywords:
- vbaac10.chm10446
ms.prod: access
api_name:
- Access.CommandButton.Hyperlink
ms.assetid: 2f5ce470-967a-450d-f661-ac1e1f370d56
ms.date: 03/01/2019
localization_priority: Normal
---


# CommandButton.Hyperlink property (Access)

You can use the **Hyperlink** property to return a reference to a **[Hyperlink](access.hyperlink.md)** object and to access the properties and methods of a control's hyperlink. Read-only.


## Syntax

_expression_.**Hyperlink**

_expression_ A variable that represents a **[CommandButton](Access.CommandButton.md)** object.


## Example

The **CreateHyperlink** procedure in the following example sets the hyperlink properties for a command button, label, or image control to the address and subaddress values passed to the procedure. The address setting is an optional argument, because a hyperlink to an object in the current database uses only the subaddress setting.

To try this example, create a form with two text box controls (**txtAddress** and **txtSubAddress**) and a command button (**cmdFollowLink**), and paste the following into the Declarations section of the form's module.

```vb
Private Sub cmdFollowLink_Click() 
 CreateHyperlink Me!cmdFollowLink, Me!txtSubAddress, _ 
 Me!txtAddress 
End Sub 
 
Sub CreateHyperlink(ctlSelected As Control, _ 
 strSubAddress As String, Optional strAddress As String) 
 Dim hlk As Hyperlink 
 Select Case ctlSelected.ControlType 
 Case acLabel, acImage, acCommandButton 
 Set hlk = ctlSelected.Hyperlink 
 With hlk 
 If Not IsMissing(strAddress) Then 
 .Address = strAddress 
 Else 
 .Address = "" 
 End If 
 .SubAddress = strSubAddress 
 .Follow 
 .Address = "" 
 .SubAddress = "" 
 End With 
 Case Else 
 MsgBox "The control '" & ctlSelected.Name _ 
 & "' does not support hyperlinks." 
 End Select 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]