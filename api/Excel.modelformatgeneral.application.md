---
title: ModelFormatGeneral.Application property (Excel)
keywords: vbaxl10.chm981073
f1_keywords:
- vbaxl10.chm981073
ms.assetid: d05c69ad-e39e-e021-b827-04a73542e816
ms.date: 05/01/2019
ms.localizationpriority: medium
---


# ModelFormatGeneral.Application property (Excel)

When used without an object qualifier, this property returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. 

When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[ModelFormatGeneral](Excel.modelformatgeneral.md)** object.


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