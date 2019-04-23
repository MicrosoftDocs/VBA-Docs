---
title: Create a DAO Recordset from a form
ms.prod: access
ms.assetid: d4bbe327-217d-ba7e-3d9f-3c89af1dcbc9
ms.date: 09/21/2018
localization_priority: Normal
---


# Create a DAO Recordset from a form

You can create a **[Recordset](../../../api/overview/Access.md)** object based on an Access form. To do so, use the **[RecordsetClone](../../../api/Access.Form.RecordsetClone.md)** property of the form. This creates a dynaset-type **Recordset** that refers to the same underlying query or data as the form. 

If a form is based on a query, referring to the **RecordsetClone** property is the equivalent of creating a dynaset with the same query. You can use the **RecordsetClone** property when you want to apply a method that cannot be used with forms, such as the **[FindFirst](../../../api/overview/Access.md)** method. The **RecordsetClone** property provides access to all the methods and properties that you can use with a dynaset.

The following example shows how to assign a **Recordset** object to the records in the Orders form.

```vb
Dim rstOrders As DAO.Recordset 
 
Set rstOrders = Forms!Orders.RecordsetClone 

```

This code always creates the type of **Recordset** being cloned (the type of **Recordset** on which the form is based); no other types are available. Note that the **Recordset** object is declared with the object library qualification. Because Access can use both DAO and ADO, it is better to fully qualify the data access variables by including the object library reference name.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]