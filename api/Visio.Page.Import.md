---
title: Page.Import Method (Visio)
keywords: vis_sdr.chm10916355
f1_keywords:
- vis_sdr.chm10916355
ms.prod: visio
api_name:
- Visio.Page.Import
ms.assetid: a84086c3-694d-8cf3-e6f7-ba84e182dd4a
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.Import Method (Visio)

Imports a file into the current document.


## Syntax

 _expression_. `Import`( `_FileName_` )

 _expression_ A variable that represents a [Page](./Visio.Page.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **String**|The name of the file to import; must be a fully qualified path.|

## Return value

 **Shape**


## Remarks

The  **Import** method imports the file specified by _FileName_ onto a page, or into a master or group.

If the path to  _FileName_ does not resolve, the **Import** method returns an error.

The file name extension indicates which import filter to use. If the filter is not installed, the  **Import** method returns an error. The **Import** method uses the default preference settings for the specified filter and does not prompt the user for non-default arguments.


## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to use the  **Import** method to import a bitmap image onto the drawing page. This example assumes that there is a file with the name _sampleImage.bmp_ on drive C of your computer.


```vb
Public Sub Import_Example() 
 
 ActivePage.Import ("C:\sampleImage.bmp") 
 
End Sub
```


