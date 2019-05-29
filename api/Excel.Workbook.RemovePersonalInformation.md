---
title: Workbook.RemovePersonalInformation property (Excel)
keywords: vbaxl10.chm199202
f1_keywords:
- vbaxl10.chm199202
ms.prod: excel
api_name:
- Excel.Workbook.RemovePersonalInformation
ms.assetid: f5cdc655-8ba9-6dd1-ab05-028d98c11972
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.RemovePersonalInformation property (Excel)

**True** if personal information can be removed from the specified workbook. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_.**RemovePersonalInformation**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Example

In this example, Microsoft Excel determines if personal information can be removed from the specified workbook and notifies the user.

```vb
Sub UsePersonalInformation() 
 
 Dim wkbOne As Workbook 
 
 Set wkbOne = Application.ActiveWorkbook 
 
 ' Determine settings and notify user. 
 If wkbOne.RemovePersonalInformation = True Then 
 MsgBox "Personal information can be removed." 
 Else 
 MsgBox "Personal information cannot be removed." 
 End If 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]