---
title: ListDataFormat.Required property (Excel)
keywords: vbaxl10.chm758082
f1_keywords:
- vbaxl10.chm758082
ms.prod: excel
api_name:
- Excel.ListDataFormat.Required
ms.assetid: ccd31ca3-906e-cacc-5ba1-48e60056d46e
ms.date: 04/30/2019
localization_priority: Normal
---


# ListDataFormat.Required property (Excel)

Returns a **Boolean** value indicating whether the schema definition of a column requires data before the row is committed. Read-only **Boolean**.


## Syntax

_expression_.**Required**

_expression_ A variable that represents a **[ListDataFormat](Excel.ListDataFormat.md)** object.


## Remarks

In Microsoft Excel, you cannot set any of the properties associated with the **ListDataFormat** object. However, you can set these properties by modifying the list on the SharePoint site.


## Example

The following example displays the setting of the **Required** property for the third column of a list on Sheet1 of the active workbook.

```vb
 
Sub Test() 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 Debug.Print objListCol.ListDataFormat.Required 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]