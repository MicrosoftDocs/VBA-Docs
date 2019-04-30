---
title: ListDataFormat.DecimalPlaces property (Excel)
keywords: vbaxl10.chm758075
f1_keywords:
- vbaxl10.chm758075
ms.prod: excel
api_name:
- Excel.ListDataFormat.DecimalPlaces
ms.assetid: 49c11876-2f79-5ca1-bdba-27e659dadc98
ms.date: 04/30/2019
localization_priority: Normal
---


# ListDataFormat.DecimalPlaces property (Excel)

Returns a **Long** value that represents the number of decimal places to show for the numbers in the **[ListColumn](Excel.ListColumn.md)** object. Read-only **Long**.


## Syntax

_expression_.**DecimalPlaces**

_expression_ A variable that represents a **[ListDataFormat](Excel.ListDataFormat.md)** object.


## Remarks

Returns 0 if the **[Type](Excel.ListDataFormat.Type.md)** setting is not appropriate for decimal places. Returns **xlAutomatic** (-4105 decimal) if the Microsoft SharePoint Foundation site is automatically determining the number of decimal places to show in the SharePoint list.

In Excel, you cannot set any of the properties associated with the **ListDataFormat** object. However, you can set these properties by modifying the list on the SharePoint site.


## Example

The following example returns the setting of the **DecimalPlaces** property for the third column of a list on Sheet1 of the active workbook.

```vb
Function GetDecimalPlaces() As Long 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns(3) 
 
 GetDecimalPlaces = objListCol.ListDataFormat.DecimalPlaces 
End Function
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]