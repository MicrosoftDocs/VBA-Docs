---
title: Shape.Application property (Excel)
keywords: vbaxl10.chm635073
f1_keywords:
- vbaxl10.chm635073
api_name:
- Excel.Shape.Application
ms.assetid: 6d86a0a4-b9ed-0ba1-22c3-30e2f100901b
ms.date: 05/14/2019
ms.localizationpriority: medium
---


# Shape.Application property (Excel)

When used without an object qualifier, this property returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. 

When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[Shape](Excel.Shape.md)** object.


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