---
title: Frameset object (Word)
keywords: vbawd10.chm2530
f1_keywords:
- vbawd10.chm2530
ms.prod: word
api_name:
- Word.Frameset
ms.assetid: d76806db-c82f-f7b6-fb85-28a649de48a7
ms.date: 06/08/2017
localization_priority: Normal
---


# Frameset object (Word)

Represents an entire frames page or a single frame on a frames page.


## Remarks

Use the  **Frameset** property of a **Document** or **Pane** object to return a **Frameset** object.


-  For properties or methods that affect all frames on a frames page, use the **Frameset** object from the **Document** object ( `ActiveWindow.Document.Frameset`).
    
- For properties or methods that affect individual frames on a frames page, use the  **Frameset** object from the **Pane** object ( `ActiveWindow.ActivePane.Frameset`).
    
This example opens a file named "Proposal.doc," creates a frames page based on the file, and adds a frame (on the left side of the page) containing a table of contents for the file.




```vb
Documents.Open "C:\My Documents\proposal.doc" 
ActiveDocument.ActiveWindow.ActivePane.NewFrameset 
ActiveDocument.ActiveWindow.ActivePane.TOCInFrameset
```

This example adds a new frame to the right of the specified frame.




```vb
ActiveDocument.ActiveWindow.ActivePane.Frameset _ 
 .AddNewFrame wdFramesetNewRight
```

This example sets the name of the third child  **Frameset** object of the frames page to "BottomFrame."




```vb
ActiveWindow.Document.Frameset _ 
 .ChildFramesetItem(3).FrameName = "BottomFrame"
```

This example links the specified frame to a local file called "Order.htm." It sets the frame to be resizable, to appear with scrollbars in a web browser, and to be 25% as high as the active window.




```vb
With ActiveDocument.ActiveWindow.ActivePane.Frameset 
 .FrameDefaultURL = "C:\My Documents\order.htm" 
 .FrameLinkToFile = True 
 .FrameResizable = True 
 .FrameScrollbarType = wdScrollbarTypeYes 
 .HeightType = wdFramesetSizeTypePercent 
 .Height = 25 
End With
```

This example sets Microsoft Word to display frame borders in the specified frames page.




```vb
ActiveDocument.ActiveWindow.ActivePane.Frameset _ 
 .FrameDisplayBorders = True
```

This example sets the frame borders on the frames page to be 6 points wide and tan.




```vb
With ActiveWindow.Document.Frameset 
 .FramesetBorderColor = wdColorTan 
 .FramesetBorderWidth = 6 
End With
```


> [!NOTE] 
> For more information on creating frames pages, see [Creating frames pages](../word/Concepts/Customizing-Word/creating-frames-pages.md).


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]