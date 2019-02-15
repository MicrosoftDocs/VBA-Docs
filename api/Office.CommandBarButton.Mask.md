---
title: CommandBarButton.Mask property (Office)
keywords: vbaof11.chm6010
f1_keywords:
- vbaof11.chm6010
ms.prod: office
api_name:
- Office.CommandBarButton.Mask
ms.assetid: de7179ac-6b39-2323-d84a-23abe3ed3167
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarButton.Mask property (Office)

Gets or sets an **IPictureDisp** object representing the mask image of a **CommandBarButton** object. The mask image determines what parts of the button image are transparent. Read/write.

> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Mask**

_expression_ A variable that represents a **[CommandBarButton](Office.CommandBarButton.md)** object.


## Remarks

When you create an image that you plan on using as a mask image, all of the areas that you want to be transparent should be white, and all of the areas that you want to show should be black.

Always set the mask after you have set the picture for a **CommandBarButton** object.


## Example

The following example sets the image and mask of the first **CommandBarButton** that the code returns. To make this work, create a mask image and a button image and substitute the paths in the sample with the paths to your images.


```vb
Sub ChangeButtonImage() 
    Dim picPicture As IPictureDisp 
    Dim picMask As IPictureDisp 
 
    Set picPicture = stdole.StdFunctions.LoadPicture( _ 
        "c:\images\picture.bmp") 
    Set picMask = stdole.StdFunctions.LoadPicture( _ 
        "c:\images\mask.bmp") 
 
    'Reference the first button on the first command bar 
    'using a With...End With block. 
    With Application.CommandBars.FindControl(msoControlButton) 
        'Change the button image. 
        .Picture = picPicture 
 
        'Use the second image to define the area of the 
        'button that should be transparent. 
        .Mask = picMask 
    End With 
End Sub
```

<br/>

The following example gets the image and mask of the first **CommandBarButton** that the code returns and outputs each of them to a file. To make this work, specify a path for the output files.

```vb
Sub GetButtonImageAndMask() 
    Dim picPicture As IPictureDisp 
    Dim picMask As IPictureDisp 
 
    With Application.CommandBars.FindControl(msoControlButton) 
        'Get the button image and mask of this CommandBarButton object 
        Set picPicture = .Picture 
        Set picMask = .Mask 
    End With 
 
    'Save the button image and mask in a folder. 
    stdole.SavePicture picPicture, "c:\image.bmp" 
    stdole.SavePicture picMask, "c:\mask.bmp" 
End Sub 

```


## See also

- [CommandBarButton object members](overview/library-reference/commandbarbutton-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]