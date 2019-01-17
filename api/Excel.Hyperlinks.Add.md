---
title: Hyperlinks.Add method (Excel)
keywords: vbaxl10.chm534073
f1_keywords:
- vbaxl10.chm534073
ms.prod: excel
api_name:
- Excel.Hyperlinks.Add
ms.assetid: 6b1299b1-c204-f0f1-c328-768c8efdb0cd
ms.date: 06/08/2017
localization_priority: Priority
---


# Hyperlinks.Add method (Excel)

Adds a hyperlink to the specified range or shape.


## Syntax

_expression_. `Add`( `_Anchor_` , `_Address_` , `_SubAddress_` , `_ScreenTip_` , `_TextToDisplay_` )

_expression_ A variable that represents a [Hyperlinks](Excel.Hyperlinks.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Anchor_|Required| **Object**|The anchor for the hyperlink. Can be either a  **[Range](Excel.Range(object).md)** or **[Shape](Excel.Shape.md)** object.|
| _Address_|Required| **String**|The address of the hyperlink.|
| _SubAddress_|Optional| **Variant**|The subaddress of the hyperlink.|
| _ScreenTip_|Optional| **Variant**|The screen tip to be displayed when the mouse pointer is paused over the hyperlink.|
| _TextToDisplay_|Optional| **Variant**|The text to be displayed for the hyperlink.|

## Return value

A  **[Hyperlink](Excel.Hyperlink.md)** object that represents the new hyperlink.


## Remarks

When you specify the  **TextToDisplay** argument, the text must be a string.


## Example

This example adds a hyperlink to cell A5.


```vb
With Worksheets(1) 
 .Hyperlinks.Add Anchor:=.Range("a5"), _ 
 Address:="https://example.microsoft.com", _ 
 ScreenTip:="Microsoft Web Site", _ 
 TextToDisplay:="Microsoft" 
End With
```

This example adds an email hyperlink to cell A5.




```vb
With Worksheets(1) 
 .Hyperlinks.Add Anchor:=.Range("a5"), _ 
 Address:="mailto:someone@example.com?subject=hello", _ 
 ScreenTip:="Write us today", _ 
 TextToDisplay:="Support" 
End With 

```


## See also


[Hyperlinks Object](Excel.Hyperlinks.md)

