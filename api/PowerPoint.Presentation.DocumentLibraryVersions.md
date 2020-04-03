---
title: Presentation.DocumentLibraryVersions property (PowerPoint)
keywords: vbapp10.chm583086
f1_keywords:
- vbapp10.chm583086
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.DocumentLibraryVersions
ms.assetid: 4c1b2055-cbbb-732d-26bd-8e6b85c26cc1
ms.date: 06/08/2017
localization_priority: Normal
---


# Presentation.DocumentLibraryVersions property (PowerPoint)

Returns a  **DocumentLibraryVersions** collection that represents the collection of versions of a shared presentation that has versioning enabled and that is stored in a document library on a server.


## Syntax

_expression_. `DocumentLibraryVersions`

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.


## Return value

DocumentLibraryVersions


## Example

The following example returns the collection of versions for the active presentation. This example assumes that the active presentation has versioning enabled and is stored in a shared document library on a server.


```vb
Dim objVersions As DocumentLibraryVersions

Set objVersions = ActivePresentation.DocumentLibraryVersions
```


## See also


[Presentation Object](PowerPoint.Presentation.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]