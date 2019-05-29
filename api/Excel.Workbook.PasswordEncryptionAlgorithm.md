---
title: Workbook.PasswordEncryptionAlgorithm property (Excel)
keywords: vbaxl10.chm199212
f1_keywords:
- vbaxl10.chm199212
ms.prod: excel
api_name:
- Excel.Workbook.PasswordEncryptionAlgorithm
ms.assetid: 2745a8da-2a61-b949-115a-7f1112a0289e
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.PasswordEncryptionAlgorithm property (Excel)

Returns a **String** indicating the algorithm that Microsoft Excel uses to encrypt passwords for the specified workbook. Read-only.


## Syntax

_expression_.**PasswordEncryptionAlgorithm**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Remarks

Use the **[SetPasswordEncryptionOptions](Excel.Workbook.SetPasswordEncryptionOptions.md)** method to specify whether Excel encrypts file properties for password-protected workbooks.


## Example

This example sets the password encryption options for the active workbook.

```vb
Sub SetPasswordOptions() 
 
 ActiveWorkbook.SetPasswordEncryptionOptions _ 
 PasswordEncryptionProvider:="Microsoft RSA SChannel Cryptographic Provider", _ 
 PasswordEncryptionAlgorithm:="RC4", _ 
 PasswordEncryptionKeyLength:=56, _ 
 PasswordEncryptionFileProperties:=True 
 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]