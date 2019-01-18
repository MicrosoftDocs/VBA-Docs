---
title: RectangularGradient object (Excel)
keywords: vbaxl10.chm856072
f1_keywords:
- vbaxl10.chm856072
ms.prod: excel
api_name:
- Excel.RectangularGradient
ms.assetid: e668d158-0436-cb27-a6f5-e27453681d66
ms.date: 06/08/2017
localization_priority: Normal
---


# RectangularGradient object (Excel)

The  **RectangularGradient** object transitions through a series of colors in a linear manner along a specific angle.


## Remarks


- Attempting to access a Gradient property of an  **Interior** object that does not have an existing gradient fill will result in a Run Time Error. Be aware of the `Interior.Pattern` property before accessing the Gradient property.
    
- If [Interior.Pattern](Excel.Interior.Pattern.md) is changed from a gradient type to a non-gradient type, the Gradient object will populate with default values.
    

 **Note**  Some things to consider when working with RectangularGradient objects


## See also



[Excel Object Model Reference](overview/Excel/object-model.md)

