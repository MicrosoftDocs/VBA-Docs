---
title: Textual criteria expressions
keywords: vbaac10.chm10884
f1_keywords:
- vbaac10.chm10884
ms.prod: access
ms.assetid: c90dbb94-daab-5ccb-4cb1-c7771d8c4fc1
ms.date: 09/21/2018
localization_priority: Normal
---


# Textual criteria expressions

To specify textual criteria for an operation, you supply a text string as part of the string expression that forms the  _criteria_ argument. This text string must be enclosed in single quotation marks (').

> [!NOTE] 
> The single quotation marks indicate to Access that the  _criteria_ argument contains a string within a string.


Suppose that you are using the ADO **[Find](../../../api/overview/Access.md)** method to find the first occurrence of a last name in an Employees table. You could construct the _criteria_ argument as in the following example, which moves the current record pointer to the first record in which an employee's last name is Buchanan. Note that the string literal `Buchanan` is enclosed in single quotation marks and the entire string comprising the criteria argument must also be enclosed in double quotation marks (").


```vb
Dim rst As New ADODB.Connection 
 
rst.open "Employees", CurrentProject.Connection,_ 
     dbOpenDynaset, adlockoptimistic) 
rst.Find "[LastName] = 'Buchanan'"
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]