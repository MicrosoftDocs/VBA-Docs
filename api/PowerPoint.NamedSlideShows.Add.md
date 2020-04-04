---
title: NamedSlideShows.Add method (PowerPoint)
keywords: vbapp10.chm515004
f1_keywords:
- vbapp10.chm515004
ms.prod: powerpoint
api_name:
- PowerPoint.NamedSlideShows.Add
ms.assetid: 413ea52c-95ba-8843-af72-952303328ebd
ms.date: 06/08/2017
localization_priority: Normal
---


# NamedSlideShows.Add method (PowerPoint)

Creates a new named slide show and adds it to the collection of named slide shows in the specified presentation. Returns a **[NamedSlideShow](PowerPoint.NamedSlideShow.md)** object that represents the new named slide show.


## Syntax

_expression_.**Add** (_Name_, _SafeArrayOfSlideIDs_)

_expression_ A variable that represents a [NamedSlideShows](PowerPoint.NamedSlideShows.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the slide show.|
| _safeArrayOfSlideIDs_|Required|**Variant**|Contains the unique slide IDs of the slides to be displayed in a slide show.|

## Return value

NamedSlideShow


## Remarks

The name you specify when you add a named slide show is the name you use as an argument to the  **[Run](PowerPoint.Application.Run.md)** method to run the named slide show.


## See also


[NamedSlideShows Object](PowerPoint.NamedSlideShows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]