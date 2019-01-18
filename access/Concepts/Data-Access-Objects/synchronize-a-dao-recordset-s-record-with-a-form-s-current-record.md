---
title: Synchronize a DAO Recordset's record with a form's current record
ms.prod: access
ms.assetid: 2960dd7d-4c60-4148-ef58-dd44f1042851
ms.date: 09/21/2018
localization_priority: Normal
---


# Synchronize a DAO Recordset's record with a form's current record

The following code example uses the **[RecordsetClone](../../../api/Access.Form.RecordsetClone.md)** property and the **[Recordset](../../../api/overview/Access.md)** object to synchronize a recordset's record with the form's current record. 

When a company name is selected from a combo box, the **[FindFirst](../../../api/overview/Access.md)** method is used to locate the record for that company, and the **Recordset** object's **[Bookmark](../../../api/overview/Access.md)** property is assigned to the form's **[Bookmark](../../../api/Access.Form.Bookmark.md)** property, causing the form to display the found record.


```vb
Sub SupplierID_AfterUpdate() 
    Dim rst As Recordset 
    Dim strSearchName As String 
 
    Set rst = Me.RecordsetClone 
    strSearchName = Str(Me!SupplierID) 
    rst.FindFirst "SupplierID = " & strSearchName 
        If rst.NoMatch Then 
            MsgBox "Record not found" 
        Else 
            Me.Bookmark = rst.Bookmark 
        End If 
    rst.Close 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]