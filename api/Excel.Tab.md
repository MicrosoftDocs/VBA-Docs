---
title: Tab Object (Excel)
keywords: vbaxl10.chm722072
f1_keywords:
- vbaxl10.chm722072
ms.prod: excel
api_name:
- Excel.Tab
ms.assetid: c6555e96-b96e-54d8-b8c6-5ab13c256d97
ms.date: 08/29/2018
---


# Tab Object (Excel)

Represents the tab of a chart or a worksheet.


## Remarks

Use the **Tab** property of the **[Chart](chart-object-excel.md)** object or **[Worksheet](worksheet-object-excel.md)** object to return a **Tab** object.

Once a  **Tab** object is returned, you can use the **[ColorIndex](Excel.Tab.ColorIndex.md)** property determine the settings of a tab for a chart or worksheet.


## Properties
|**Name**|
|:-----|
|[Application Property](tab-application-property-excel.md)|
|[Color Property](tab-color-property-excel.md)|
|[ColorIndex Property](tab-colorindex-property-excel.md)|
|[Creator Property](tab-creator-property-excel.md)|
|[Parent Property](tab-parent-property-excel.md)|
|[ThemeColor Property](tab-themecolor-property-excel.md)|
|[TintAndShade Property](tab-tintandshade-property-excel.md)|

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

[Chart.Tab Property](chart-tab-property-excel.md)
[Worksheet.Tab Property](worksheet-tab-property-excel.md)
[Excel Object Model Reference](./overview/Excel/object-model.md)

