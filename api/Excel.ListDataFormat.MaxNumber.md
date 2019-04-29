---
title: ListDataFormat.MaxNumber property (Excel)
keywords: vbaxl10.chm758080
f1_keywords:
- vbaxl10.chm758080
ms.prod: excel
api_name:
- Excel.ListDataFormat.MaxNumber
ms.assetid: 61262a29-7a35-e351-71fa-0b217285e2b3
ms.date: 04/30/2019
localization_priority: Normal
---


# ListDataFormat.MaxNumber property (Excel)

Returns a **Variant** containing the maximum value allowed in this field in the list column. Read-only **Variant**.


## Syntax

_expression_.**MaxNumber**

_expression_ A variable that represents a **[ListDataFormat](Excel.ListDataFormat.md)** object.


## Remarks

The **Nothing** object is returned if a maximum value number has not been specified or if the **Type** property setting is such that a maximum value for the column is not applicable.

This property is used only for lists that are linked to a SharePoint site.

In Microsoft Excel, you cannot set any of the properties associated with the **ListDataFormat** object. However, you can set these properties by modifying the list on the SharePoint site.


## Example

The following example displays the setting of the **MaxNumber** property for the third column of a list on Sheet1 of the active workbook.

```vb
 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 Debug.Print objListCol.ListDataFormat.MaxNumber
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]