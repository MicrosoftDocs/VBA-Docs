---
title: ListDataFormat.IsPercent property (Excel)
keywords: vbaxl10.chm758077
f1_keywords:
- vbaxl10.chm758077
ms.prod: excel
api_name:
- Excel.ListDataFormat.IsPercent
ms.assetid: 34154cf9-358a-0db9-4b93-fe3b3f2b8dce
ms.date: 04/30/2019
localization_priority: Normal
---


# ListDataFormat.IsPercent property (Excel)

Returns a **Boolean** value. Returns **True** only if the number data for the **[ListColumn](Excel.ListColumn.md)** object will be shown in percentage formatting. Read-only **Boolean**. 


## Syntax

_expression_.**IsPercent**

_expression_ A variable that represents a **[ListDataFormat](Excel.ListDataFormat.md)** object.


## Remarks

This property is used only for lists that are linked to a Microsoft SharePoint Foundation site.

In Excel, you cannot set any of the properties associated with the **ListDataFormat** object. However, you can set these properties by modifying the list on the SharePoint site.


## Example

The following example returns the setting of the **IsPercent** property for the third column of the list on Sheet1 of the active workbook.

```vb
Function GetIsPercent() As Boolean 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 GetIsPercent = objListCol.ListDataFormat.IsPercent 
End Function
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]