---
title: Range.AllowEdit property (Excel)
keywords: vbaxl10.chm144239
f1_keywords:
- vbaxl10.chm144239
api_name:
- Excel.Range.AllowEdit
ms.assetid: 9f03054c-190f-ce3b-54db-bc6e19b7e1c6
ms.date: 05/10/2019
ms.localizationpriority: medium
---


# Range.AllowEdit property (Excel)

Returns a **Boolean** value that indicates if the range can be edited on a protected worksheet.


## Syntax

_expression_.**AllowEdit**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Example

In this example, Microsoft Excel notifies the user whether cell A1 can be edited on a protected worksheet.

```vb
Sub UseAllowEdit() 
 
 Dim wksOne As Worksheet 
 
 Set wksOne = Application.ActiveSheet 
 
 ' Protect the worksheet 
 wksOne.Protect 
 
 ' Notify the user about editing cell A1. 
 If wksOne.Range("A1").AllowEdit = True Then 
 MsgBox "Cell A1 can be edited." 
 Else 
 Msgbox "Cell A1 cannot be edited." 
 End If 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
