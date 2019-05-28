---
title: Workbook.Protect method (Excel)
keywords: vbaxl10.chm199217
f1_keywords:
- vbaxl10.chm199217
ms.prod: excel
api_name:
- Excel.Workbook.Protect
ms.assetid: 0e270b93-7b0b-cc68-c7c0-4002024f4292
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.Protect method (Excel)

Protects a workbook so that it cannot be modified.


## Syntax

_expression_.**Protect** (_Password_, _Structure_, _Windows_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Password_|Optional| **Variant**|A string that specifies a case-sensitive password for the worksheet or workbook. If this argument is omitted, you can unprotect the worksheet or workbook without using a password. Otherwise, you must specify the password to unprotect the worksheet or workbook. If you forget the password, you cannot unprotect the worksheet or workbook.<br/><br/>Use strong passwords that combine uppercase and lowercase letters, numbers, and symbols. Weak passwords don't mix these elements. Strong password: Y6dh!et5. Weak password: House27.<br/><br/>Passwords should be 8 or more characters in length. A pass phrase that uses 14 or more characters is better.<br/><br/>It is critical that you remember your password. If you forget your password, Microsoft cannot retrieve it. Store the passwords that you write down in a secure place away from the information that they help protect. |
| _Structure_|Optional| **Variant**| **True** to protect the structure of the workbook (the relative position of the sheets). The default value is **False**.|
| _Windows_|Optional| **Variant**| **True** to protect the workbook windows. If this argument is omitted, the windows aren't protected.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
