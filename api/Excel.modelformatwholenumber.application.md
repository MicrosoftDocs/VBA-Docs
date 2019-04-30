---
title: ModelFormatWholeNumber.Application property (Excel)
keywords: vbaxl10.chm987073
f1_keywords:
- vbaxl10.chm987073
ms.assetid: a2dada26-38f7-967a-be69-3f75f911c05e
ms.date: 05/01/2019
ms.prod: excel
localization_priority: Normal
---


# ModelFormatWholeNumber.Application property (Excel)

When used without an object qualifier, this property returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. 

When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[ModelFormatWholeNumber](Excel.modelformatwholenumber.md)** object.


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