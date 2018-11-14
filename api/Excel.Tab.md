---
title: Tab object (Excel)
keywords: vbaxl10.chm722072
f1_keywords:
- vbaxl10.chm722072
ms.prod: excel
api_name:
- Excel.Tab
ms.assetid: c6555e96-b96e-54d8-b8c6-5ab13c256d97
ms.date: 09/05/2018
---


# Tab object (Excel)

Represents the tab of a chart or a worksheet.


## Remarks

Use the **Tab** property of the **[Chart](Excel.Chart(object).md)** object or **[Worksheet](Excel.Worksheet.md)** object to return a **Tab** object.

Once a  **Tab** object is returned, you can use the **[ColorIndex](Excel.Tab.ColorIndex.md)** property determine the settings of a tab for a chart or worksheet.


## Properties

|**Name**|
|:-----|
|[Application property](Excel.Tab.Application.md)|
|[Color property](Excel.Tab.Color.md)|
|[ColorIndex property](Excel.Tab.ColorIndex.md)|
|[Creator property](Excel.Tab.Creator.md)|
|[Parent property](Excel.Tab.Parent.md)|
|[ThemeColor property](Excel.Tab.ThemeColor.md)|
|[TintAndShade property](Excel.Tab.TintAndShade.md)|

## Example

In the following example, Microsoft Excel determines if the worksheet's first tab color index is set to none and notifies the user.


```vb
Sub CheckTab() 
 
 ' Determine if color index of 1st tab is set to none. 
 If Worksheets(1).Tab.ColorIndex = xlColorIndexNone Then 
 MsgBox "The color index is set to none for the first " & _ 
 "worksheet tab." 
 Else 
 MsgBox "The color index for the tab of the first worksheet " & _ 
 "is not set none." 
 End If 
 
End Sub
```

## See also

- [Chart.Tab property](Excel.Chart.Tab.md)
- [Worksheet.Tab property](Excel.Worksheet.Tab.md)
- [Excel object model reference](overview/Excel/object-model.md)

