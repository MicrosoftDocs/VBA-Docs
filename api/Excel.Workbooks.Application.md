---
title: Workbooks.Application property (Excel)
keywords: vbaxl10.chm202073
f1_keywords:
- vbaxl10.chm202073
api_name:
- Excel.Workbooks.Application
ms.assetid: 20988e9c-58e6-4e08-ba67-e3611eb111b2
ms.date: 05/18/2019
ms.localizationpriority: medium
---


# Workbooks.Application property (Excel)

When used without an object qualifier, this property returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. 

When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[Workbooks](Excel.Workbooks.md)** object.


## Example

This example displays a message about the application that created _myObject_.

```vb
Set myObject = ActiveWorkbook 
If myObject.Application.Value = "Microsoft Excel" Then 
 MsgBox "This is an Excel Application object." 
Else 
 MsgBox "This is not an Excel Application object." 
End If
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]