---
title: Creator property (Excel Graph)
keywords: vbagr10.chm65685
f1_keywords:
- vbagr10.chm65685
ms.prod: excel
api_name:
- Excel.Creator
ms.assetid: 79d72908-f141-1d3a-d8db-c10db7b33537
ms.date: 04/10/2019
localization_priority: Normal
---


# Creator property (Excel Graph)

Returns a 32-bit integer that indicates the application in which the specified object was created. If the object was created in Graph, this property returns the string MSGR, which is equivalent to the hexadecimal number 4D534752. Read-only **[XlCreator](excel.xlcreator.md)**.

## Syntax

_expression_.**Creator**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Example

This example displays a message about the creator of _myChart_.

```vb
If myChart.Creator = &h4D534752 Then 
    MsgBox "This is a Graph object" 
Else 
    MsgBox "This is not a Graph object" 
End If
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]