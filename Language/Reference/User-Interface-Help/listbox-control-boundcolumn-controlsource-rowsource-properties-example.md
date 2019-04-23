---
title: ListBox control, BoundColumn, ControlSource, RowSource properties example
keywords: fm20.chm5225190
f1_keywords:
- fm20.chm5225190
ms.prod: office
ms.assetid: 5aa015d5-0a2b-ca93-940c-3faf4dd9d900
ms.date: 11/14/2018
localization_priority: Normal
---


# ListBox control, BoundColumn, ControlSource, RowSource properties example

The following example uses a range of worksheet cells in a **[ListBox](listbox-control.md)** and, when the user selects a row from the list, displays the row index in another worksheet cell. This code sample uses the **[RowSource](rowsource-property.md)**, **[BoundColumn](boundcolumn-property.md)**, and **[ControlSource](controlsource-property.md)** properties.

To use this example:

- Copy this sample code to the Declarations portion of a form. 
- Make sure that the form contains a **ListBox** named ListBox1. 
- In the worksheet, enter data in cells A1:E4. 
- Make sure cell A6 contains no data.


```vb
Private Sub UserForm_Initialize() 
 
ListBox1.ColumnCount = 5 
ListBox1.RowSource = "a1:e4" 
 
ListBox1.ControlSource = "a6" 
'Place the ListIndex into cell a6 
ListBox1.BoundColumn = 0 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
