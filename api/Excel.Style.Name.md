---
title: Style.Name property (Excel)
keywords: vbaxl10.chm177090
f1_keywords:
- vbaxl10.chm177090
ms.prod: excel
api_name:
- Excel.Style.Name
ms.assetid: 4ad63465-afe0-fc96-3ec7-62318d8ac1e2
ms.date: 05/16/2019
localization_priority: Normal
---


# Style.Name property (Excel)

Returns a **String** value that represents the name of the object.


## Syntax

_expression_.**Name**

_expression_ A variable that represents a **[Style](Excel.Style.md)** object.


## Example

This example displays the name of style one in the active workbook, first in the language of the macro and then in the language of the user.

```vb
With ActiveWorkbook.Styles(1) 
 MsgBox "The name of the style: " & .Name 
 MsgBox "The localized name of the style: " & .NameLocal 
End With
```

<br/>

The following example displays the name of the default **ListObject** object in Sheet1 of the active workbook.

```vb
 
Sub Test 
 Dim wrksht As Worksheet 
 Dim oListObj As ListObject 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set oListObj = wrksht.ListObjects(1) 
 
 MsgBox oListObj.Name 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]