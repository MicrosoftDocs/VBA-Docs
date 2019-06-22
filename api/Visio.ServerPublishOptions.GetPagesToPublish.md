---
title: ServerPublishOptions.GetPagesToPublish method (Visio)
keywords: vis_sdr.chm17962380
f1_keywords:
- vis_sdr.chm17962380
ms.prod: visio
api_name:
- Visio.ServerPublishOptions.GetPagesToPublish
ms.assetid: e5dacddd-9b3d-7d18-afff-82ee6a042b03
ms.date: 06/08/2017
localization_priority: Normal
---


# ServerPublishOptions.GetPagesToPublish method (Visio)

Returns an array of pages that are set to be published to a server.


## Syntax

_expression_. `GetPagesToPublish`( `_Flags_` , `_PublishPages_` , `_NamesArray()_` )

_expression_ A variable that represents a **[ServerPublishOptions](Visio.ServerPublishOptions.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Flags_|Required| **[VisLangFlags](Visio.VisLangFlags.md)**|Out parameter. Indicates whether universal or local page names are returned in  _NamesArray_. See Remarks for possible values.|
| _PublishPages_|Required| **[VisPublishPages](Visio.VisPublishPages.md)**|Out parameter. Indicates whether all pages or selected pages are set to be published. See Remarks for possible values.|
| _NamesArray()_|Required| **String**|Out parameter. Returns the names of the pages set to be published.|

## Return value

 **Nothing**


## Remarks

The  _Flags_ parameter can be one of the following **VisLangFlags** constants.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visLangLocal**|0|The page name is a local name.|
| **visLangUniversal**|1|The page name is a universal name.|

The  _PublishPages_ parameter can be one of the following **VisPublishPages** constants.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visPublishPageAll**|0|Publish all pages.|
| **visPublishPageSelect**|1|Publish selected pages.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]