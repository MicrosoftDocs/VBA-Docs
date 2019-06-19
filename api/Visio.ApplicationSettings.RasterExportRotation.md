---
title: ApplicationSettings.RasterExportRotation property (Visio)
keywords: vis_sdr.chm16262545
f1_keywords:
- vis_sdr.chm16262545
ms.prod: visio
api_name:
- Visio.ApplicationSettings.RasterExportRotation
ms.assetid: 660b22ff-11b6-bfaf-1949-18e5e9c57d64
ms.date: 06/08/2017
localization_priority: Normal
---


# ApplicationSettings.RasterExportRotation property (Visio)

Determines the rotation that is applied to the exported image when you call the  **Export** method of the **[Master](Visio.Master.md)**, **[Page](Visio.Page.md)**, **[Selection](Visio.Selection.md)**, or **[Shape](Visio.Shape.md)** object to export the specified object to a BMP, GIF, JPG, PNG, or TIFF file. Read/write.


## Syntax

_expression_.**RasterExportRotation**

 _expression_ An expression that returns an **[ApplicationSettings](Visio.ApplicationSettings.md)** object.


## Return value

 **[VisRasterExportRotation](Visio.VisRasterExportRotation.md)**


## Remarks

The value of the  **RasterExportRotation** property must be one of the following **VisRasterExportRotation** constants.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visRasterNoRotation**|0|No rotation, the default.|
| **visRasterRotateLeft**|1|Rotate left.|
| **visRasterRotateRight**|2|Rotate right.|

For any given session of Microsoft Visio, when the  **RasterExportRotation** property value is set, either programmatically or in the user interface, the setting then becomes the new default for the remainder of the session. However, it is not persisted to the next session.

The setting of the  **RasterExportRotation** property corresponds to the style of rotation selected in the **Rotation** list in the **Output Options** dialog box for the corresponding file type in the Microsoft Visio user interface. (Click the **File** tab, click **Save As**, in the  **Save as type** list, select the file type, and then click **Save**.)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]