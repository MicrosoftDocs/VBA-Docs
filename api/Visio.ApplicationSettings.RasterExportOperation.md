---
title: ApplicationSettings.RasterExportOperation property (Visio)
keywords: vis_sdr.chm16262540
f1_keywords:
- vis_sdr.chm16262540
ms.prod: visio
api_name:
- Visio.ApplicationSettings.RasterExportOperation
ms.assetid: 7f53b4a6-6497-01ca-2219-575065d4c9f4
ms.date: 06/08/2017
localization_priority: Normal
---


# ApplicationSettings.RasterExportOperation property (Visio)

Determines the export operation that is applied to the exported image when you call the  **Export** method of the **[Master](Visio.Master.md)**, **[Page](Visio.Page.md)**, **[Selection](Visio.Selection.md)**, or **[Shape](Visio.Shape.md)** object to export the specified object to a JPG file. Read/write.


## Syntax

_expression_.**RasterExportOperation**

 _expression_ An expression that returns an **[ApplicationSettings](Visio.ApplicationSettings.md)** object.


## Return value

 **[VisRasterExportOperation](Visio.VisRasterExportOperation.md)**


## Remarks

The value of the  **RasterExportOperation** property must be one of the following **VisRasterExportOperation** constants.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visRasterBaseline**|0|Baseline operation, the default.|
| **visRasterProgressive**|1|Progressive operation.|

For any given session of Microsoft Visio, when the  **RasterExportOperation** property value is set, either programmatically or in the user interface, the setting then becomes the new default for the remainder of the session. However, it is not persisted to the next session.

The setting of the  **RasterExportOperation** property corresponds to the **Operation** setting in the **JPG Output Options** dialog box in the Microsoft Visio user interface. (Click the **File** tab, click **Save As**, in the  **Save as type** list, select **JPEG File Interchange Format (*.jpg)**, and then click  **Save**.)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]