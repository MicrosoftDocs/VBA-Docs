---
title: PivotItem.SourceNameStandard property (Excel)
keywords: vbaxl10.chm246093
f1_keywords:
- vbaxl10.chm246093
ms.prod: excel
api_name:
- Excel.PivotItem.SourceNameStandard
ms.assetid: f8e25ad0-7a97-c19c-85b5-bf25e3553ca8
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotItem.SourceNameStandard property (Excel)

Returns a **String** that represents the PivotTable items' source name in standard English (United States) format settings. Read-only.


## Syntax

_expression_.**SourceNameStandard**

_expression_ A variable that represents a **[PivotItem](Excel.PivotItem.md)** object.


## Remarks

This property is used when an item has a localized version and its **SourceNameStandard** property value differs from the **[SourceName](Excel.PivotItem.SourceName.md)** property value, such as with date formatting.


## Example

This example displays the source name for the sixth item on the fifth field of the active PivotTable. The example assumes that a PivotTable exists on the active worksheet and that the data source contains at least five fields and six items per field.

```vb
Sub CheckSourceNameStandard() 
 
 Dim pvtTable As PivotTable 
 Dim pvtField As PivotField 
 Dim pvtItem As PivotItem 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtField = pvtTable.PivotFields(5) 
 Set pvtItem = pvtField.PivotItems(6) 
 
 ' Display source name. 
 MsgBox "The source name is: " & pvtItem.SourceNameStandard 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]