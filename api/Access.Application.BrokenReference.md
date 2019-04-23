---
title: Application.BrokenReference property (Access)
keywords: vbaac10.chm12593
f1_keywords:
- vbaac10.chm12593
ms.prod: access
api_name:
- Access.Application.BrokenReference
ms.assetid: 20a55f4b-5fe4-9231-bbef-e90c66f88b90
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.BrokenReference property (Access)

Returns a **Boolean** indicating whether the current database has any broken references to databases or type libraries. **True** if there are any broken references. Read-only.


## Syntax

_expression_.**BrokenReference**

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Remarks

To test the validity of a specific reference, use the **[IsBroken](Access.Reference.IsBroken.md)** property of the **[Reference](Access.Reference.md)** object.


## Example

This example checks to see if there are any broken references in the current database and reports the results to the user.

```vb
' Looping variable. 
Dim refLoop As Reference 
' Output variable. 
Dim strReport As String 
 
' Test whether there are broken references. 
If Application.BrokenReference = True Then 
 strReport = "The following references are broken:" & vbCr 
 
 ' Test validity of each reference. 
 For Each refLoop In Application.References 
 If refLoop.IsBroken = True Then 
 strReport = strReport & " " & refLoop.Name & vbCr 
 End If 
 Next refLoop 
Else 
 strReport = "All references in the current database are valid." 
End If 
 
' Display results. 
MsgBox strReport
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]