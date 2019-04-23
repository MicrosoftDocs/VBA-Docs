---
title: Report.Error event (Access)
keywords: vbaac10.chm13880
f1_keywords:
- vbaac10.chm13880
ms.prod: access
api_name:
- Access.Report.Error
ms.assetid: 06d88711-df19-6453-a7ce-095d3d02674f
ms.date: 03/08/2019
localization_priority: Normal
---


# Report.Error event (Access)

The **Error** event occurs when a run-time error is produced in Microsoft Access when a report has the focus.


## Syntax

_expression_.**Error** (_DataErr_, _Response_)

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DataErr_|Required|**Integer**|The error code returned by the **Err** object when an error occurs. You can use the _DataErr_ argument with the **Error** function to map the number to the corresponding error message. |
| _Response_|Required|**Integer**|The setting determines whether or not an error message is displayed. The _Response_ argument can be one of the following intrinsic constants: <ul><li><b>acDataErrContinue</b>  Ignore the error and continue without displaying the default Microsoft Access error message. You can supply a custom error message in place of the default error message.</li><li><b>acDataErrDisplay</b>  (Default) Display the default Access error message.</li></ul>|

## Return value

Nothing


## Remarks

This includes Access database engine errors, but not run-time errors in Visual Basic or errors from ADO.

To run a macro or event procedure when this event occurs, set the **[OnError](access.report.onerror.md)** property to the name of the macro or to [Event Procedure].

By running an event procedure or a macro when an **Error** event occurs, you can intercept an Access error message and display a custom message that conveys a more specific meaning for your application.
  

## Example

The following example shows how you can replace a default error message with a custom error message. When Access returns an error message indicating that it has found a duplicate key (error code 3022), this event procedure displays a message that gives more application-specific information to users.

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




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]