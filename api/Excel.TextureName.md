---
title: TextureName property (Excel Graph)
keywords: vbagr10.chm5208045
f1_keywords:
- vbagr10.chm5208045
ms.prod: excel
api_name:
- Excel.TextureName
ms.assetid: a2c0e2af-5f16-f181-0404-49223de24a97
ms.date: 04/12/2019
localization_priority: Normal
---


# TextureName property (Excel Graph)

Returns the name of the custom texture file for the specified fill. Read-only **String**.

## Syntax

_expression_.**TextureName**

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Remarks

This property is read-only. Use the **[UserPicture](excel.userpicture.md)** or **[UserTextured](excel.usertextured.md)** method to set the texture file for the fill.

## Example

This example changes the user-defined texture type for the chart's fill format.

```vb
With myChart.ChartArea.Fill 
 If .Type = msoFillTextured Then 
 If .TextureType = msoTextureUserDefined Then 
 If .TextureName = "brick.bmp" Then 
 .UserTextured "stone.bmp" 
 End If 
 End If 
 End If 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]