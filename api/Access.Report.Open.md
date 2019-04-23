---
title: Report.Open event (Access)
keywords: vbaac10.chm13876
f1_keywords:
- vbaac10.chm13876
ms.prod: access
api_name:
- Access.Report.Open
ms.assetid: d170b67d-3123-6f51-6cf8-38433736f104
ms.date: 03/08/2019
localization_priority: Normal
---


# Report.Open event (Access)

The **Open** event occurs before a report is previewed or printed.


## Syntax

_expression_.**Open** (_Cancel_)

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required|**Integer**|The setting determines if the opening of the form or report occurs. Setting the _Cancel_ argument to **True** (1) cancels the opening of the form or report.|

## Return value

Nothing

## Remarks

For example, an **Open** macro or event procedure can open a custom dialog box in which the user enters the criteria to filter the set of records to display on a form or the date range to include for a report.

When you open a report based on an underlying query, Microsoft Access runs the **Open** macro or event procedure before it runs the underlying query for the report. This enables the user to specify criteria for the report before it opens; for example, in a custom dialog box you display when the **Open** event occurs.

If your application can have more than one form loaded at a time, use the **Activate** and **Deactivate** events instead of the **Open** event to display and hide custom toolbars when the focus moves to a different form.

When the **Close** event occurs, you can open another window or request the user's name to make a log entry indicating who used the form or report.

If you are trying to decide whether to use the **Open** or **Load** event for your macro or event procedure, one significant difference is that the **Open** event can be canceled, but the **Load** event can't. For example, if you are dynamically building a record source for a form in an event procedure for the form's **Open** event, you can cancel opening the form if there are no records to display. Similarly, the **Unload** event can be canceled, but the **Close** event can't.


## Example

The following example shows how to use a Structured Query Language (SQL) statement to establish the data source of a report as it is opened.

```vb
Private Sub Report_Open(Cancel As Integer)

    On Error GoTo Error_Handler

    Me.Caption = "My Application"

    DoCmd.OpenForm FormName:="frmReportSelector_MemberList", _
    Windowmode:=acDialog

    'Cancel the report if "cancel" was selected on the dialog form.

    If Forms!frmReportSelector_MemberList!txtContinue = "no" Then
        Cancel = True
        GoTo Exit_Procedure
    End If
    Me.RecordSource = ReplaceWhereClause(Me.RecordSource, _
      Forms!frmReportSelector_MemberList!txtWhereClause)

Exit_Procedure:
    Exit Sub

Error_Handler:
    MsgBox Err.Number & ": " & Err.Description
    Resume Exit_Procedure
    Resume

End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]