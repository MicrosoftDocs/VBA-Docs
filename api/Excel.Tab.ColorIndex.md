---
title: Tab.ColorIndex property (Excel)
keywords: vbaxl10.chm723074
f1_keywords:
- vbaxl10.chm723074
ms.prod: excel
api_name:
- Excel.Tab.ColorIndex
ms.assetid: 4c257c58-613e-dbc9-095f-3609feffe64c
ms.date: 05/17/2019
localization_priority: Normal
---


# Tab.ColorIndex property (Excel)

Returns or sets a **Variant** value that represents the color of the specified worksheet tab or chart tab.


## Syntax

_expression_.**ColorIndex**

_expression_ A variable that represents a **[Tab](Excel.Tab.md)** object.


## Remarks

After a **Tab** object is returned, you can use the **ColorIndex** property to determine the settings of a tab for a chart or worksheet.

The color is specified as an index value in the current color palette from 1 to 56 or the **[XlColorIndex](Excel.XlColorIndex.md)** value **xlColorIndexNone**.

## Example

In the following example, Microsoft Excel determines whether the first worksheet's tab color index is set to none and notifies the user.

```vb
Sub CheckTab() 
 
 ' Determine if color index of 1st tab is set to none. 
 If Worksheets(1).Tab.ColorIndex = xlColorIndexNone Then 
  MsgBox "The color index is set to none for the first " & _ 
  "worksheet tab." 
 Else 
  MsgBox "The color index for the tab of the first worksheet " & _ 
  "is not set to none." 
 End If 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]