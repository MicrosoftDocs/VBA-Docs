---
title: SoundNote.Application property (Excel)
keywords: vbaxl10.chm257073
f1_keywords:
- vbaxl10.chm257073
ms.assetid: 3adf2c05-3fc5-6a29-8c4f-ea6021db2802
ms.date: 05/16/2019
ms.prod: excel
localization_priority: Normal
---


# SoundNote.Application property (Excel)

> [!NOTE] 
> This property should not be used. Sound notes have been removed from Microsoft Excel.

When used without an object qualifier, this property returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. 

When used with an object qualifier, this property returns an **Application** object that represents the creator of the specified object (you can use this property with an OLE Automation object to return the application of that object). Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **SoundNote** object.


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