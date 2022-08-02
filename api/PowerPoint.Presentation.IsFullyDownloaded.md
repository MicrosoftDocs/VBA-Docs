---
title: Presentation.IsFullyDownloaded property (PowerPoint)
keywords: vbapp10.chm583138
f1_keywords:
- vbapp10.chm583138
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.IsFullyDownloaded
ms.date: 07/27/2022
ms.author: ononder
ms.localizationpriority: medium
---


# Presentation.IsFullyDownloaded property (PowerPoint)

**True** if the presentation has finished downloading all of the content. Read-only **Boolean.**

## Syntax

_expression_.**IsFullyDownloaded**

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.

## Remarks

When you open a presentation with content that is large in size, PowerPoint may serve the document in parts as partial documents. This allows for documents to be opened, edited, and collaborated upon quickly, while the larger media parts (e.g., videos), continue to load in the background. Similarly, since media is handled separately from the rest of the document, collaboration is smoother when media is inserted during a collaboration session.

Because certain content can be deferred initially, some actions can't be taken on that content until the deferred content (e.g., video) is loaded. Additionally, there are certain actions like Save As, Export to Video, etc. that won’t function until all the deferred content are downloaded. User initiated operations will display UI informing the user of download progress, but that’s not possible for programmatic operations.  If you programmatically attempt to call an API to execute an action in these cases, it will fail.

To understand this state programmatically, you may query **Presentation.IsFullyDownloaded** property before calling any of the impacted APIs and add error handling to capture the failure and retry the operation once the presentation is fully downloaded.

## Example

The following example displays a message indicating if the active presentation is fully downloaded or not.


```vb
If ActivePresentation.IsFullyDownloaded Then
    MsgBox "Everything is downloaded"
Else
    MsgBox "Not fully downloaded"
End If
```


## See also

[Presentation Object](PowerPoint.Presentation.md)

[Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]