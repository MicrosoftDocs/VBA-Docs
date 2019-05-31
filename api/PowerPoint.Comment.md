---
title: Comment object (PowerPoint)
keywords: vbapp10.chm642000
f1_keywords:
- vbapp10.chm642000
ms.prod: powerpoint
api_name:
- PowerPoint.Comment
ms.assetid: c1071b54-eeaa-0cec-13f0-b635da9511d8
ms.date: 06/08/2017
localization_priority: Normal
---


# Comment object (PowerPoint)

Represents a comment on a given slide or slide range. The  **Comment** object is a member of the **[Comments](PowerPoint.Comments.md)** collection object.


## Remarks

Use the following properties to access comment data:


|||
|:-----|:-----|
|[Author](PowerPoint.Comment.Author.md)|The author's full name|
|[AuthorIndex](PowerPoint.Comment.AuthorIndex.md)|The author's index in the list of comments|
|[AuthorInitials](PowerPoint.Comment.AuthorInitials.md)|The author's initials|
|[DateTime](PowerPoint.Comment.DateTime.md)|The date and time the comment was created|
|[Text](PowerPoint.Comment.Text.md)|The text of the comment|
|[Left](PowerPoint.Comment.Left.md), [Top](PowerPoint.Comment.Top.md)|The comment's screen coordinates|

## Example

Use  **[Comments](PowerPoint.Slide.Comments.md)** (_index_), where _index_ is the number of the comment, or the **[Item](PowerPoint.Comments.Item.md)** method to access a single comment on a slide. This example displays the author of the first comment on the first slide. If there are no comments, it displays a message stating such.


```vb
Sub ShowComment()

    With ActivePresentation.Slides(1).Comments

        If .Count > 0 Then

            MsgBox "The first comment on this slide is by " & .Item(1).Author

        Else

            MsgBox "There are no comments on this slide."

        End If

    End With

End Sub
```

This example displays a message containing the author, date and time, and contents of all the messages on the first slide.




```vb
Sub SlideComments()

    Dim cmtExisting As Comment
    Dim cmtAll As Comments
    Dim strComments As String

    Set cmtAll = ActivePresentation.Slides(1).Comments

    If cmtAll.Count > 0 Then
        For Each cmtExisting In cmtAll
            strComments = strComments & cmtExisting.Author & vbTab & _
                cmtExisting.DateTime & vbTab & cmtExisting.Text & vbLf
        Next
        MsgBox "The comments in your document are as follows:" & vbLf & strComments
    Else
        MsgBox "This slide doesn't have any comments."
    End If

End Sub
```


## See also


[PowerPoint Object Model Reference](overview/PowerPoint/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]