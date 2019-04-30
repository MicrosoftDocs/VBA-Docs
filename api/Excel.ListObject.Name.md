---
title: ListObject.Name property (Excel)
keywords: vbaxl10.chm734088
f1_keywords:
- vbaxl10.chm734088
ms.prod: excel
api_name:
- Excel.ListObject.Name
ms.assetid: fbbdf2f9-6c5f-6ebe-35b1-74aab63971a4
ms.date: 04/30/2019
localization_priority: Normal
---


# ListObject.Name property (Excel)

Returns or sets a **String** value that represents the name of the **ListObject** object.


## Syntax

_expression_.**Name**

_expression_ An expression that returns a **[ListObject](Excel.ListObject.md)** object.


## Return value

String


## Remarks

This name is used solely as a unique identifier for the **[Item](Excel.ListObjects.Item.md)** property of the **ListObjects** collection objects. This property can only be set through the object model.

By default, each **ListObject** object name begins with the word "List", followed by a number (no spaces). If an attempt is made to set the **Name** property to a name already used by another **ListObject** object, a run-time error is thrown.


## Example

The following example displays the name of the default **ListObject** object on Sheet1 of the active workbook.

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
