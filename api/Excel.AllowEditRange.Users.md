---
title: AllowEditRange.Users property (Excel)
keywords: vbaxl10.chm725078
f1_keywords:
- vbaxl10.chm725078
ms.prod: excel
api_name:
- Excel.AllowEditRange.Users
ms.assetid: 71f3c7ed-2fba-d97b-e443-674836e6bddb
ms.date: 04/04/2019
localization_priority: Normal
---


# AllowEditRange.Users property (Excel)

Returns a **[UserAccessList](Excel.UserAccessList.md)** object for the protected range on a worksheet.


## Syntax

_expression_.**Users**

_expression_ A variable that represents an **[AllowEditRange](Excel.AllowEditRange.md)** object.


## Example

In this example, Microsoft Excel displays the name of the first user allowed access to the first protected range on the active worksheet. This example assumes that a range has been chosen to be protected and that a particular user has been given access to this range.

```vb
Sub DisplayUserName() 
 
 Dim wksSheet As Worksheet 
 
 Set wksSheet = Application.ActiveSheet 
 
 ' Display name of user with access to protected range. 
 MsgBox wksSheet.Protection.AllowEditRanges(1).Users(1).Name 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]