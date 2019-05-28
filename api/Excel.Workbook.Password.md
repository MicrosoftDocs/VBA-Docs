---
title: Workbook.Password property (Excel)
keywords: vbaxl10.chm199209
f1_keywords:
- vbaxl10.chm199209
ms.prod: excel
api_name:
- Excel.Workbook.Password
ms.assetid: 5eaaf8cd-4344-946e-ecfa-c0f48946d2f2
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.Password property (Excel)

Returns or sets the password that must be supplied to open the specified workbook. Read/write **String**.


## Syntax

_expression_.**Password**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Remarks

Use strong passwords that combine uppercase and lowercase letters, numbers, and symbols. Weak passwords don't mix these elements. 

- Strong password: Y6dh!et5
- Weak password: House27

Passwords should be 8 or more characters in length. A pass phrase that uses 14 or more characters is better. It is critical that you remember your password. If you forget your password, Microsoft cannot retrieve it. Store the passwords that you write down in a secure place away from the information that they help protect. 


## Example

In this example, Microsoft Excel opens a workbook named Password.xls, sets a password for it, and then closes the workbook. This example assumes that a file named Password.xls exists on the C:\ drive.

```vb
Sub UsePassword() 
 
 Dim wkbOne As Workbook 
 
 Set wkbOne = Application.Workbooks.Open("C:\Password.xls") 
 
 wkbOne.Password = InputBox ("Enter Password") 
 wkbOne.Close 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
