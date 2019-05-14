---
title: ShapeNode.Application property (Excel)
ms.prod: excel
api_name:
- Excel.ShapeNode.Application
ms.assetid: 98e77d56-875c-7696-2b2d-5f36409fa129
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeNode.Application property (Excel)

When used without an object qualifier, this property returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. 

When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[ShapeNode](Excel.ShapeNode.md)** object.


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