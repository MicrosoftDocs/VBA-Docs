---
title: HeadersFooters.Header property (PowerPoint)
keywords: vbapp10.chm542005
f1_keywords:
- vbapp10.chm542005
ms.prod: powerpoint
api_name:
- PowerPoint.HeadersFooters.Header
ms.assetid: 83748bf7-10a4-1ee7-4eef-4ef8fd38b7da
ms.date: 06/08/2017
localization_priority: Normal
---


# HeadersFooters.Header property (PowerPoint)

Returns a  **[HeaderFooter](PowerPoint.HeaderFooter.md)** object that represents the header that appears at the top of a slide or in the upper-left corner of a notes page, handout, or outline. Read-only.


## Syntax

_expression_. `Header`

_expression_ A variable that represents a [HeadersFooters](PowerPoint.HeadersFooters.md) object.


## Return value

HeaderFooter


## Example

This example sets the header text for the handout master for the active presentation. This text will appear in the upper-left corner of the page when you print your presentation as an outline or a handout.


```vb
Set myHandHF = Application.ActivePresentation.HandoutMaster _
    .HeadersFooters

myHandHF.Header.Text = "Third Quarter Report"
```


## See also


[HeadersFooters Object](PowerPoint.HeadersFooters.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]