---
title: Close a report automatically if it does not contain any records
ms.prod: access
ms.assetid: 9b160bd3-6eca-f907-ae5b-4327c3c1618e
ms.date: 09/26/2018
localization_priority: Normal
---


# Close a report automatically if it does not contain any records

The following example shows how to use the **[NoData](../../../api/Access.Report.NoData.md)** event to cancel opening or printing a report when it has no data. A message box notifying the user that the report has no data is also displayed.


```vb
Private Sub Report_NoData (Cancel As Integer) 
     
    ' Display message to user. 
    MsgBox "There are no records to report", vbExclamation, "No Records" 
 
    ' Close the report. 
    Cancel = True 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]