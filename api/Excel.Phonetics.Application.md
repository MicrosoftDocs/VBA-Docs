---
title: Phonetics.Application property (Excel)
keywords: vbaxl10.chm657073
f1_keywords:
- vbaxl10.chm657073
ms.prod: excel
api_name:
- Excel.Phonetics.Application
ms.assetid: 346c0218-d19f-beaf-a192-09c1fa95da25
ms.date: 06/08/2017
localization_priority: Normal
---


# Phonetics.Application property (Excel)

When used without an object qualifier, this property returns an  **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

_expression_. `Application`

_expression_ A variable that represents a [Phonetics](Excel.Phonetics.md) object.


## Example

This example displays a message about the application that created  `myObject`.


```vb
Set myObject = ActiveWorkbook 
If myObject.Application.Value = "Microsoft Excel" Then 
 MsgBox "This is an Excel Application object." 
Else 
 MsgBox "This is not an Excel Application object." 
End If
```


## See also


[Phonetics Object](Excel.Phonetics.md)

