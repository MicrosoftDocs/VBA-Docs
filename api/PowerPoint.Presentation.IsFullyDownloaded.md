---
title: Presentation.IsFullyDownloaded property (PowerPoint)
keywords: vbapp10.chm583138
f1_keywords:
- vbapp10.chm583138
api_name:
- PowerPoint.Presentation.IsFullyDownloaded
ms.date: 08/02/2022
ms.author: ononder
ms.localizationpriority: medium
---


# Presentation.IsFullyDownloaded property (PowerPoint)

**True** if the presentation has finished downloading all of the content. Read-only **Boolean.**

## Syntax

_expression_.**IsFullyDownloaded**

_expression_ A variable that represents a [Presentation](PowerPoint.Presentation.md) object.

## Remarks

When you open a presentation with content that is large in size, PowerPoint may serve the document in parts as partial documents. This allows you to open, edit, and collaborate on documents quickly, while the larger media parts (e.g., videos), continue to load in the background. Similarly, since media is handled separately from the rest of the document, collaboration is smoother when media is inserted during a collaboration session.

Because certain content can be deferred initially, some actions can't be taken until the deferred content is loaded. Additionally, there are certain actions like Save As, Export to Video, etc. that won’t function until all the deferred content are downloaded. If you initiate one of these operations, PowerPoint will display UI informing you of the download progress, but that’s not possible for programmatic operations. If you programmatically attempt to call an API to execute an action while content is still downloading, it will fail.

To understand this state programmatically, you can query **Presentation.IsFullyDownloaded** property before you call any of the impacted APIs. Add error handling to capture any failures and retry the operation once the presentation is fully downloaded.

## Example

The following example displays a message indicating if the active presentation is fully downloaded or not.


```vb
If ActivePresentation.IsFullyDownloaded Then
    MsgBox "Presentation is downloaded."
Else
    MsgBox "PowerPoint is still downloading the presentation."
End If
```


## See also

[Presentation Object](PowerPoint.Presentation.md)

[Work with Partial Documents](~/powerpoint/How-to/work-with-partial-documents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]