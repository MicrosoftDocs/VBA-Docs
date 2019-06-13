---
title: Selection.Type property (Publisher)
keywords: vbapb10.chm851971
f1_keywords:
- vbapb10.chm851971
ms.prod: publisher
api_name:
- Publisher.Selection.Type
ms.assetid: 4dfcfecc-dd76-36b6-21df-34c3865b3064
ms.date: 06/13/2019
localization_priority: Normal
---


# Selection.Type property (Publisher)

Returns a **[PbSelectionType](publisher.pbselectiontype.md)** constant that represents the selection type. Read-only.


## Syntax

_expression_.**Type**

_expression_ A variable that represents a **[Selection](Publisher.Selection.md)** object.


## Remarks

The **Type** property value can be one of the **PbSelectionType** constants.


## Example

This example checks to see if the selection is text, and if it is, makes the selected text bold.

```vb
Sub IfCellData() 
 Dim rowTable As Row 
 If Selection.Type = pbSelectionText Then 
 Selection.TextRange.Font.Bold = msoTrue 
 End If 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]