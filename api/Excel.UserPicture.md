---
title: UserPicture method (Excel Graph)
keywords: vbagr10.chm67165
f1_keywords:
- vbagr10.chm67165
ms.prod: excel
api_name:
- Excel.UserPicture
ms.assetid: ad8e3079-c063-2bb6-e462-11a0e8ecfba6
ms.date: 04/09/2019
localization_priority: Normal
---


# UserPicture method (Excel Graph)

Fills the specified shape with an image.

## Syntax

_expression_.**UserPicture** (_PictureFile_, _PictureFormat_, _PictureStackUnit_, _PicturePlacement_)

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_PictureFile_ |Required |**Variant**|The name of the specified picture file.|
|_PictureFormat_ |Optional|**[XlChartPictureType](excel.xlchartpicturetype.md)**|The format of the specified picture. Can be one of the **XlChartPictureType** constants.|
|_PictureStackUnit_ |Optional |**Variant**|The stack or scale unit for the specified picture (depends on the _PictureFormat_ argument).|
|_PicturePlacement_ |Optional|**[XlChartPicturePlacement](excel.xlchartpictureplacement.md)**|The placement of the specified picture. Can be one of the **XlChartPicturePlacement** constants.|

## Example

This example sets the chart's fill format so that it's based on a user-supplied picture.

```vb
With myChart.ChartArea.Fill 
 .UserPicture PictureFile:="C:\My Documents\brick.bmp" 
 .Visible = True 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]