---
title: NegativeBarFormat.ColorType Property (Excel)
keywords: vbaxl10.chm887073
f1_keywords:
- vbaxl10.chm887073
ms.prod: excel
api_name:
- Excel.NegativeBarFormat.ColorType
ms.assetid: 01485eab-0aa3-278e-2976-02e0d0757a4f
ms.date: 06/08/2017
---


# NegativeBarFormat.ColorType Property (Excel)

 Specifies whether to use the same fill color as positive data bars. Read/write


## Syntax

 _expression_ . **ColorType**

 _expression_ A variable that represents a **[NegativeBarFormat](Excel.NegativeBarFormat.md)** object.


### Return Value

 **[XlDataBarNegativeColorType](Excel.XlDataBarNegativeColorType.md)**


## Remarks

If you set the  **ColorType** property to **xlDataBarColor** , use the **[Color](Excel.NegativeBarFormat.Color.md)** property of the **NegativeBarFormat** object to return a **[FormatColor](Excel.FormatColor.md)** object that you can use to specify the fill color.


## See also


#### Concepts


[NegativeBarFormat Object](Excel.NegativeBarFormat.md)

