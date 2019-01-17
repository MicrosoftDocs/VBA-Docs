---
title: Workbook.BuiltinDocumentProperties property (Excel)
keywords: vbaxl10.chm199081
f1_keywords:
- vbaxl10.chm199081
ms.prod: excel
api_name:
- Excel.Workbook.BuiltinDocumentProperties
ms.assetid: 3efffd7d-0681-ecbc-000a-b71eceb3f92a
ms.date: 06/08/2017
localization_priority: Priority
---


# Workbook.BuiltinDocumentProperties property (Excel)

Returns a  **[DocumentProperties](Office.DocumentProperties.md)** collection that represents all the built-in document properties for the specified workbook. Read-only.


## Syntax

_expression_. `BuiltinDocumentProperties`

_expression_ A variable that represents a [Workbook](./Excel.Workbook.md) object.


## Remarks

This property returns the entire collection of built-in document properties. Use the  **Item** method to return a single member of the collection (a **[DocumentProperty](Office.DocumentProperty.md)** object) by specifying either the name of the property or the collection index (as a number).

You can refer to document properties either by index value or by name. The following list shows the available built-in document property names:



|Title Subject Author Keywords Comments Template Last Author Revision Number Application Name Last Print Date|Creation Date Last Save Time Total Editing Time Number of Pages Number of Words Number of Characters Security Category Format Manager|Company Number of Bytes Number of Lines Number of Paragraphs Number of Slides Number of Notes Number of Hidden Slides Number of Multimedia Clips Hyperlink Base Number of Characters (with spaces)|

Container applications aren't required to define values for every built-in document property. If Microsoft Excel doesn't define a value for one of the built-in document properties, reading the  **Value** property for that document property causes an error.

Because the  **Item** method is the default method for the **DocumentProperties** collection, the following statements are identical:




```vb
BuiltinDocumentProperties.Item(1) 
BuiltinDocumentProperties(1)
```

Use the  **[CustomDocumentProperties](Excel.Workbook.CustomDocumentProperties.md)** property to return the collection of custom document properties.


## Example

This example displays the names of the built-in document properties as a list on worksheet one.


```vb
rw = 1 
Worksheets(1).Activate 
For Each p In ActiveWorkbook.BuiltinDocumentProperties 
    Cells(rw, 1).Value = p.Name 
    rw = rw + 1 
Next
```


## See also


[Workbook Object](Excel.Workbook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]