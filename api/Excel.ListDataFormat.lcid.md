---
title: ListDataFormat.lcid property (Excel)
keywords: vbaxl10.chm758078
f1_keywords:
- vbaxl10.chm758078
ms.prod: excel
api_name:
- Excel.ListDataFormat.lcid
ms.assetid: 954812f2-d50e-8eff-429d-37da5cd8cff1
ms.date: 04/30/2019
localization_priority: Normal
---


# ListDataFormat.lcid property (Excel)

Returns a **Long** value that represents the LCID for the **[ListColumn](Excel.ListColumn.md)** object that is specified in the schema definition. Read-only **Long**.


## Syntax

_expression_.**lcid**

_expression_ A variable that represents a **[ListDataFormat](Excel.ListDataFormat.md)** object.


## Remarks

In Microsoft Excel, the LCID indicates the currency symbol to be used when this is an **xlListDataTypeCurrency** type (**[XlListDataType](excel.xllistdatatype.md)** enumeration). Returns 0 (which is the Language Neutral LCID) when no locale is set for the data type of the column.

This property is used only for tables that are linked to a Microsoft SharePoint Foundation site.

In Excel, you cannot set any of the properties associated with the **ListDataFormat** object. However, you can set these properties by modifying the list on the SharePoint site.


## Example

The following example displays the setting of the **lcid** property for the third column of the list on Sheet1 of the active workbook.

```vb
Sub DisplayLCID() 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 MsgBox "List LCID: " & objListCol.ListDataFormat.lcid 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]