---
title: ServerPublishOptions.SetPagesToPublish method (Visio)
keywords: vis_sdr.chm17962375
f1_keywords:
- vis_sdr.chm17962375
ms.prod: visio
api_name:
- Visio.ServerPublishOptions.SetPagesToPublish
ms.assetid: 9d874876-e053-d6fb-04c2-8e162a0457ec
ms.date: 06/08/2017
localization_priority: Normal
---


# ServerPublishOptions.SetPagesToPublish method (Visio)

Specifies the pages to publish to a server.


## Syntax

_expression_. `SetPagesToPublish`( `_PublishPages_` , `_NamesArray()_` , `_ Flags_` )

_expression_ A variable that represents a **[ServerPublishOptions](Visio.ServerPublishOptions.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PublishPages_|Required| **[VisPublishPages](Visio.VisPublishPages.md)**|Indicates whether all pages or selected pages are to be published. See Remarks for possible values.|
| _NamesArray()_|Required| **String**|The names of the pages to be published, if  _PublishPages_ is **visPublishPageSelect**.|
| _Flags_|Required| **[VisLangFlags](Visio.VisLangFlags.md)**|Indicates whether universal or local page names are specified in  _NamesArray_. See Remarks for possible values.|

## Return value

 **Nothing**


## Remarks

The  _PublishPages_ parameter must be one of the following **VisPublishPages** constants.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visPublishPageAll**|0|Publish all pages.|
| **visPublishPageSelect**|1|Publish selected pages.|

The  _Flags_ parameter must be one of the following **VisLangFlags** constants.



|Constant|Value|Description|
|:-----|:-----|:-----|
| **visLangLocal**|0|The page name is a local name.|
| **visLangUniversal**|1|The page name is a universal name.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]