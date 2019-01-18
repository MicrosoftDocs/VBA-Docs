---
title: Filter a report using a form's filter
ms.prod: access
ms.assetid: 2b029c13-5abd-4865-cd05-25d094a97b9f
ms.date: 09/26/2018
localization_priority: Normal
---


# Filter a report using a form's filter

The following example illustrates how to open a report based on the filtered contents of a form. To do this, specify the form's **[Filter](../../../api/Access.Form.Filter(property).md)** property as the value of the **[OpenReport](../../../api/Access.DoCmd.OpenReport.md)** method's _WhereCondition_ argument.


```vb
Private Sub cmdOpenReport_Click() 
    If Me.Filter = "" Then 
        MsgBox "Apply a filter to the form first." 
    Else 
        DoCmd.OpenReport "rptCustomers", acViewReport, , Me.Filter 
    End If 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]