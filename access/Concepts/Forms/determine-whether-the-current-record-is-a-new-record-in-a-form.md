---
title: Determine whether the current record is a new record in a form
ms.prod: access
ms.assetid: 04aa27cd-b6b1-1397-c177-bac939780492
ms.date: 09/25/2018
localization_priority: Normal
---


# Determine whether the current record is a new record in a form

The following example shows how to use the **NewRecord** property to determine if the current record is a new record. The **NewRecordMark** procedure sets the current record to the variable _intnewrec_. If the record is new, a message notifies the user. You could call this procedure when the [Current](../../../api/Access.Form.Current.md) event for a form occurs.


```vb
Sub NewRecordMark(frm As Form) 
    Dim intnewrec As Integer 
 
    intnewrec = frm.NewRecord 
    If intnewrec = True Then 
    MsgBox "You're in a new record." _ 
        & "@Do you want to add new data?" _ 
        & "@If not, move to an existing record." 
    End If 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
