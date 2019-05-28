---
title: Workbook.OpenLinks method (Excel)
keywords: vbaxl10.chm199120
f1_keywords:
- vbaxl10.chm199120
ms.prod: excel
api_name:
- Excel.Workbook.OpenLinks
ms.assetid: cae33bab-892e-0861-e4ec-8a334097e0d1
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.OpenLinks method (Excel)

Opens the supporting documents for a link or links.


## Syntax

_expression_.**OpenLinks** (_Name_, _ReadOnly_, _Type_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the Microsoft Excel or DDE/OLE link, as returned from the **[LinkSources](Excel.Workbook.LinkSources.md)** method.|
| _ReadOnly_|Optional| **Variant**| **True** to open documents as read-only. The default value is **False**.|
| _Type_|Optional| **Variant**|One of the constants of **[XlLink](Excel.XlLink.md)** that specifies the link type.|

## Example

This example opens OLE link one in the active workbook.

```vb
linkArray = ActiveWorkbook.LinkSources(xlOLELinks) 
ActiveWorkbook.OpenLinks linkArray(1)
```

<br/>

This example opens all supporting Microsoft Excel documents for the active workbook.

```vb
Sub OpenAllLinks() 
 Dim arLinks As Variant 
 Dim intIndex As Integer 
 arLinks = ActiveWorkbook.LinkSources(xlExcelLinks) 
 
 If Not IsEmpty(arLinks) Then 
 For intIndex = LBound(arLinks) To UBound(arLinks) 
 ActiveWorkbook.OpenLinks arLinks(intIndex) 
 Next intIndex 
 Else 
 MsgBox "The active workbook contains no external links." 
 End If 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]