---
title: WebNavigationBarHyperlinks.Add method (Publisher)
keywords: vbapb10.chm8585220
f1_keywords:
- vbapb10.chm8585220
ms.prod: publisher
api_name:
- Publisher.WebNavigationBarHyperlinks.Add
ms.assetid: 6cd0c43a-fec1-c9b8-dc86-00e1cc314087
ms.date: 06/18/2019
localization_priority: Normal
---


# WebNavigationBarHyperlinks.Add method (Publisher)

Adds a new **[Hyperlink](publisher.hyperlink.md)** object to the specified **WebNavigationBarHyperlinks** collection and returns the new **Hyperlink** object. 


## Syntax

_expression_.**Add** (_Address_, _RelativePage_, _PageID_, _TextToDisplay_, _Index_)

_expression_ A variable that represents a **[WebNavigationBarHyperlinks](publisher.webnavigationbarhyperlinks.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Address_|Optional| **String**|The address of the new hyperlink. If _RelativePage_ is **pbHlinkTargetTypeURL** (default) or **pbHlinkTargetTypeEmail**, _Address_ must be specified, or an error occurs.|
|_RelativePage_|Optional| **[PbHlinkTargetType](Publisher.PbHlinkTargetType.md)**|The type of hyperlink to add. Can be one of the **PbHlinkTargetType** constants. The default is **pbHlinkTargetTypeURL**.|
|_PageID_|Optional| **Long**|The page ID of the destination page for the new hyperlink. If _RelativePage_ is **pbHlinkTargetTypePageID**, _PageID_ must be specified, or an error occurs. The page ID corresponds to the **[PageID](Publisher.Page.PageID.md)** property of the destination page.|
|_TextToDisplay_|Optional| **String**|The display text of the new hyperlink. |
|_Index_|Optional| **Long**|The index of the new **Hyperlink** object in the **WebNavigationBarHyperlinks** collection.|

## Return value

Hyperlink



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]