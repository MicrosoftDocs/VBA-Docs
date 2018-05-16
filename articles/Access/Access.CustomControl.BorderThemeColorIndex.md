---
title: CustomControl.BorderThemeColorIndex Property (Access)
keywords: vbaac10.chm14634
f1_keywords:
- vbaac10.chm14634
ms.prod: access
api_name:
- Access.CustomControl.BorderThemeColorIndex
ms.assetid: 4e4d6aeb-dd68-b16f-375a-be4c3cf95286
ms.date: 06/08/2017
---


# CustomControl.BorderThemeColorIndex Property (Access)

Gets or sets a value that represents a color in the applied color theme associated with the  **BorderColor** property of the specified object. Read/write **Long**.


## Syntax

 _expression_. **BorderThemeColorIndex**

 _expression_ A variable that represents a **CustomControl** object.


## Remarks

The  **BorderThemeColorIndex** property contains one of the index values listed in the following table.



|**Index Value**|**Description**|
|:-----|:-----|
|0|Text 1|
|1|Background 1|
|2|Text 2|
|3|Background 2|
|4|Accent 1|
|5|Accent 2|
|6|Accent 3|
|7|Accent 4|
|8|Accent 5|
|9|Accent 6|
|10|Hyperlink|
|11|Followed Hyperlink|
If no theme is applied, the  **BorderThemeColorIndex** property contains -1.

This property is not surfaced in the property sheet.


## Example

The following code example sets the Border Color to the Text 2 color by setting the  **BorderThemeColorIndex** property.


```vb
Me.ctl.BorderThemeColorIndex=2
```


## See also


#### Concepts


[CustomControl Object](Access.CustomControl.md)

