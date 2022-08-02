---
title: Work with Partial Documents
keywords: vbapp10.chm583138
f1_keywords:
- vbapp10.chm583138
ms.prod: powerpoint
ms.date: 07/27/2022
ms.author: ononder
ms.localizationpriority: medium
---


# Work with Partial Documents

When you open a presentation with content that is large in size, PowerPoint may serve the document in parts as partial documents. This allows for documents to be opened, edited, and collaborated upon quickly, while the larger media parts (e.g., videos), continue to load in the background. Similarly, since media is handled separately from the rest of the document, collaboration is smoother when media is inserted during a collaboration session.

Because certain content can be deferred initially, some actions can't be taken on that content until the deferred content (e.g., video) is loaded. Additionally, there are certain actions like Save As, Export to Video, etc. that won’t function until all the deferred content are downloaded. User initiated operations will display UI informing the user of download progress, but that’s not possible for programmatic operations.  If you programmatically attempt to call an API to execute an action in these cases, it will fail with the following error message.


```
Run-time error '-2147188128 (80048260)':
<object> (unknown member) : This method isn't supported until the presentation is fully downloaded. Visit this URL for more information: https://go.microsoft.com/fwlink/?linkid=2172228
```


## Understanding the Fully Downloaded State

To understand if a presentation is fully downloaded programmatically, you may query [Presentation.IsFullyDownloaded](~/api/PowerPoint.IsFullyDownloaded.md) property before calling any of the impacted APIs.


```vb
If ActivePresentation.IsFullyDownloaded Then
    MsgBox "Everything is downloaded"
Else
    MsgBox "Not fully downloaded"
End If
```


## Handling Errors

 You may also add some error handling to capture the failure and retry the operation once the presentation is fully downloaded. If the error value is `-2147188128` or `0x80048260`, the operation has failed because the presentation is not fully downloaded.
**Err.Number** can be used as a key to identify these failures.


```vb
Sub TestCopySlide()
    On Error GoTo eh    
    ActivePresentation.Slides(1).Copy    
    Exit Sub
eh:
    If Err.Number = -2147188128 Then
        MsgBox "Cannot copy because the presentation is not fully downloaded"
    Else
        MsgBox "Failed due to reason other than incomplete parts: " & Err.Description
    End If
    Debug.Print Err.Number, Err.Description
End Sub
```


## Impacted APIs

The following is a list of impacted OM API calls which may return the error code:

|Name|
|:-----|
|[Presentation.Export](~/api/PowerPoint.Presentation.Export.md)|
|[Presentation.ExportAsFixedFormat](~/api/PowerPoint.Presentation.ExportAsFixedFormat.md)|
|[Presentation.ExportAsFixedFormat2](~/api/PowerPoint.Presentation.ExportAsFixedFormat2.md)|
|[Presentation.SaveAs](~/api/PowerPoint.Presentation.SaveAs.md)|
|[Presentation.SaveCopyAs](~/api/PowerPoint.Presentation.SaveCopyAs.md)|
|[Presentation.SaveCopyAs2](~/api/PowerPoint.Presentation.SaveCopyAs2.md)|
|[Presentation.Password](~/api/PowerPoint.Presentation.Password.md)|
|[Presentation.WritePassword](~/api/PowerPoint.Presentation.WritePassword.md)|
|[Selection.Copy](~/api/PowerPoint.Selection.Copy.md)|
|[Selection.Cut](~/api/PowerPoint.Selection.Cut.md)|
|[Shape.Copy](~/api/PowerPoint.Shape.Copy.md)|
|[Shape.Cut](~/api/PowerPoint.Shape.Cut.md)|
|[ShapeRange.Cut](~/api/PowerPoint.ShapeRange.Cut.md)|
|[ShapeRange.Copy](~/api/PowerPoint.ShapeRange.Copy.md)|
|[Shapes.Paste](~/api/PowerPoint.Shapes.Paste.md)|
|[Slide.Copy](~/api/PowerPoint.Slide.Copy.md)|
|[Slide.Cut](~/api/PowerPoint.Slide.Cut.md)|
|[Slide.Export](~/api/PowerPoint.Slide.Export.md)|
|[SlideRange.Copy](~/api/PowerPoint.SlideRange.Copy.md)|
|[SlideRange.Cut](~/api/PowerPoint.SlideRange.Cut.md)|
|[SlideRange.Export](~/api/PowerPoint.SlideRange.Export.md)|
|[Slides.Paste](~/api/PowerPoint.Slides.Paste.md)|
|[CustomLayouts.Paste](~/api/PowerPoint.CustomLayouts.Paste.md)|
|[View.Paste](~/api/PowerPoint.View.Paste.md)|
|[View.PasteSpecial](~/api/PowerPoint.View.PasteSpecial.md)|
|[MediaFormat.Resample](~/api/PowerPoint.MediaFormat.Resample.md)|
|[MediaFormat.ResampleFromProfile](~/api/PowerPoint.MediaFormat.ResampleFromProfile.md)|
|[Player.Play](~/api/PowerPoint.Player.Play.md)|


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]