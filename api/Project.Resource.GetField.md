---
title: Resource.GetField method (Project)
ms.prod: project-server
api_name:
- Project.Resource.GetField
ms.assetid: 36fbbc13-272e-72f4-ebbe-2c13f67abbe7
ms.date: 06/08/2017
localization_priority: Normal
---


# Resource.GetField method (Project)

Returns the value of the specified resource custom field.


## Syntax

_expression_. `GetField`( `_FieldID_` )

_expression_ A variable that represents a [Resource](./Project.Resource.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**Long**|For a local custom field, can be one of the **[PjField](Project.PjField.md)** constants for resource custom fields. For an enterprise custom field, use the **[FieldNameToFieldConstant](Project.Application.FieldNameToFieldConstant.md)** method to get the FieldID.|

## Return value

 **String**


## Example

The following example displays the value of a local resource custom field specified by the user.


```vb
Sub DisplayField() 
    Dim Temp As String 
 
    Temp = InputBox$("Enter the name of the field you want to see:") 
    Temp = LCase(Temp) 
 
    Select Case Temp 
        Case "name" 
            MsgBox (ActiveCell.Resource.GetField(FieldID:=pjResourceName)) 
        Case "initials" 
            MsgBox (ActiveCell.Resource.GetField(FieldID:=pjResourceInitials)) 
        Case "standard rate" 
            MsgBox (ActiveCell.Resource.GetField(FieldID:=pjResourceStandardRate)) 
        Case "" 
            End 
        Case Else 
            MsgBox "You entered an invalid field. Please try again." 
            End 
    End Select 
End Sub
```

For an example that uses an enterprise resource custom field, see the **[SetField](Project.Resource.SetField.md)** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]