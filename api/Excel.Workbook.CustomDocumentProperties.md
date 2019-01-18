---
title: Workbook.CustomDocumentProperties property (Excel)
keywords: vbaxl10.chm199094
f1_keywords:
- vbaxl10.chm199094
ms.prod: excel
api_name:
- Excel.Workbook.CustomDocumentProperties
ms.assetid: 8470adbb-5b10-96ba-71f7-c667c33b6707
ms.date: 06/08/2017
localization_priority: Priority
---


# Workbook.CustomDocumentProperties property (Excel)

Returns or sets a  **[DocumentProperties](Office.DocumentProperties.md)** collection that represents all the custom document properties for the specified workbook.


## Syntax

_expression_. `CustomDocumentProperties`

_expression_ A variable that represents a [Workbook](./Excel.Workbook.md) object.


## Remarks

This property returns the entire collection of custom document properties. Use the  **Item** method to return a single member of the collection (a **DocumentProperty** object) by specifying either the name of the property or the collection index (as a number).

Because the  **Item** method is the default method for the **DocumentProperties** collection, the following statements are identical:

 `CustomDocumentProperties.Item("Complete")`

 `CustomDocumentProperties("Complete")`

Use the  **[BuiltinDocumentProperties](Excel.Workbook.BuiltinDocumentProperties.md)** property to return the collection of built-in document properties.

Properties of type  **msoPropertyTypeString** cannot exceed 255 characters in length.


## Example

This example displays the names and values of the custom document properties as a list on worksheet one.


```vb
rw = 1 
Worksheets(1).Activate 
For Each p In ActiveWorkbook.CustomDocumentProperties 
    Cells(rw, 1).Value = p.Name 
    Cells(rw, 2).Value = p.Value 
    rw = rw + 1 
Next
```


## See also


[Workbook Object](Excel.Workbook.md)

