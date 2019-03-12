---
title: Form.Error event (Access)
keywords: vbaac10.chm13658
f1_keywords:
- vbaac10.chm13658
ms.prod: access
api_name:
- Access.Form.Error
ms.assetid: ed8229fb-4169-8be5-dc2e-a543ca3bfff3
ms.date: 03/08/2019
localization_priority: Normal
---


# Form.Error event (Access)

The **Error** event occurs when a run-time error is produced in Microsoft Access when a form has the focus.


## Syntax

_expression_.**Error** (_DataErr_, _Response_)

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DataErr_|Required|**Integer**|The error code returned by the **Err** object when an error occurs. You can use the _DataErr_ argument with the **Error** function to map the number to the corresponding error message. |
| _Response_|Required|**Integer**|The setting determines whether an error message is displayed. The _Response_ argument can be one of the following intrinsic constants.<ul><li><p><b>acDataErrContinue</b>  Ignore the error and continue without displaying the default Microsoft Access error message. You can supply a custom error message in place of the default error message.</p></li><li><p><b>acDataErrDisplay</b>  (Default) Display the default Access error message.</p></li></ul>|

## Remarks

This includes Access database engine errors, but not run-time errors in Visual Basic or errors from ADO.

To run a macro or event procedure when this event occurs, set the **OnError** property to the name of the macro or to [Event Procedure].

By running an event procedure or a macro when an **Error** event occurs, you can intercept an Access error message and display a custom message that conveys a more specific meaning for your application.
    

## Example

The following example shows how you can replace a default error message with a custom error message. When Access returns an error message indicating it has found a duplicate key (error code 3022), this event procedure displays a message that gives more application-specific information to users.

To try the example, add the following event procedure to a form that is based on a table with a unique employee ID number as the key for each record.

```vb
Private Sub Form_Error(DataErr As Integer, Response As Integer) 
    Const conDuplicateKey = 3022 
    Dim strMsg As String 
 
    If DataErr = conDuplicateKey Then 
        Response = acDataErrContinue 
        strMsg = "Each employee record must have a unique " _ 
            & "employee ID number. Please recheck your data." 
        MsgBox strMsg 
    End If 
End Sub
```

<br/>

The following example shows how you can replace a default error message with a custom error message.

```vb
Private Sub Form_Error(DataErr As Integer, Response As Integer)
    Select Case DataErr
        Case 2113
            MsgBox "Only numbers are acceptable in this box", vbCritical, "Call 1-800-123-4567"
            Response = acDataErrContinue
        Case 2237
            MsgBox "You can only choose from the dropdown box"
            Response = acDataErrContinue
        Case 3022
            MsgBox "You entered a value that exists already in another record"
            Response = acDataErrContinue
            SSN.Value = SSN.OldValue
        Case 3314
            MsgBox "The DOH is required, so you cannot leave this field empty"
            Response = acDataErrContinue
        Case Else
            Response = acDataErrDisplay
    End Select
    ActiveControl.Undo
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
