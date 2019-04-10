---
title: Model3DFormat.ResetModel method (PowerPoint)
keywords: vbapp10.chm743020
f1_keywords:
- vbapp10.chm743020
ms.prod: powerpoint
api_name:
- PowerPoint.Model3DFormat.ResetModel
ms.date: 04/11/2019
localization_priority: Normal
---


# Model3DFormat.ResetModel method (PowerPoint)

Changes the rotation of the specified shape around the x-axis by the specified number of degrees. 


## Syntax

_expression_.**ResetModel** (_ResetSize_)

_expression_ A variable that represents a **[Model3DFormat](PowerPoint.Model3DFormat.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ResetSize_|Required|**Boolean**|**True** to reset the 3D model frame to the same size as when a model is first inserted; **False** to leave the 3D model frame size alone.|

## Remarks

Use the **ResetModel** method to restore 3D model properties back to default settings.  Any camera settings, shape properties, light properties, and animation properties are set to the same values that are applied when a 3D model is first inserted into a document.  

The size of the 3D model frame can also be conditionally changed if the parameter _ResetFrameSize_ is set to **True**.


## Example

This example resets the properties of a 3D model on _myDocument_ back to the settings that the model had immediately after being first inserted into a document, and also resets the frame size to default dimensions.

```vb
Set myDocument = ActivePresentation.Slides(1)

myDocument.Shapes(1).Model3D.ResetModel(True)
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]