---
title: Application.DoCmd property (Access)
keywords: vbaac10.chm12511
f1_keywords:
- vbaac10.chm12511
ms.prod: access
api_name:
- Access.Application.DoCmd
ms.assetid: 171fb56a-b39f-4439-e841-ae4bbbd71719
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.DoCmd property (Access)

You can use the **DoCmd** property to access the read-only **[DoCmd](Access.DoCmd.md)** object and its related methods. Read-only **DoCmd**.


## Syntax

_expression_.**DoCmd**

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Example

The following example opens a form in Form view and moves to a new record.


```vb
Sub ShowNewRecord() 
 DoCmd.OpenForm "Employees", acNormal 
 DoCmd.GoToRecord , , acNewRec 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
