---
title: QueryTables.Application property (Excel)
keywords: vbaxl10.chm520073
f1_keywords:
- vbaxl10.chm520073
api_name:
- Excel.QueryTables.Application
ms.assetid: 428b9fbb-8119-0de6-d86a-719edbe5d667
ms.date: 05/03/2019
ms.localizationpriority: medium
---


# QueryTables.Application property (Excel)

When used without an object qualifier, this property returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. 

When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[QueryTables](Excel.QueryTables.md)** object.


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