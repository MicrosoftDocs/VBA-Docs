---
title: ServerPublishOptions.IncludePage method (Visio)
keywords: vis_sdr.chm17962365
f1_keywords:
- vis_sdr.chm17962365
ms.prod: visio
api_name:
- Visio.ServerPublishOptions.IncludePage
ms.assetid: 6af3f654-3b08-a990-8f0c-b05bb046a0b4
ms.date: 06/08/2017
localization_priority: Normal
---


# ServerPublishOptions.IncludePage method (Visio)

Includes the specified page for publication when the document is published as a VDW file.


## Syntax

_expression_. `IncludePage`( `_PageName_` , `_Flags_` )

_expression_ A variable that represents a **[ServerPublishOptions](Visio.ServerPublishOptions.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PageName_|Required| **String**|The name of the page to be published.|
| _Flags_|Required| **[VisLangFlags](Visio.VisLangFlags.md)**|Indicates whether a universal or local page name is specified in PageName. See Remarks for possible values.|

## Return value

 **Nothing**


## Remarks

The  _Flags_ parameter must be one of the following **VisLangFlags** constants.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visLangLocal**|0|The page name is a local name.|
| **visLangUniversal**|1|The page name is a universal name.|

Calling the  **IncludePage** method corresponds to selecting a page in the **Pages** list in the **Publish Settings** dialog box (click the **File** tab, click **Save & Send**, click  **Save to SharePoint**, click  **Web Drawing (*.vdw)**, click  **Save As**, and then click  **Options**).

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]