---
title: Hyperlink.Application property (Excel)
keywords: vbaxl10.chm535073
f1_keywords:
- vbaxl10.chm535073
api_name:
- Excel.Hyperlink.Application
ms.assetid: 14b44cb6-712b-23a0-c215-ef624b83d38b
ms.date: 04/26/2019
ms.localizationpriority: medium
---


# Hyperlink.Application property (Excel)

When used without an object qualifier, this property returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. 

When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[Hyperlink](Excel.Hyperlink.md)** object.


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