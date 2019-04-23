---
title: CommandBarButton.Picture property (Office)
keywords: vbaof11.chm6009
f1_keywords:
- vbaof11.chm6009
ms.prod: office
api_name:
- Office.CommandBarButton.Picture
ms.assetid: b9a2d133-23a8-ac09-8b8b-08eda1210717
ms.date: 01/03/2019
localization_priority: Normal
---


# CommandBarButton.Picture property (Office)

Gets or sets an **IPictureDisp** object representing the image of a **CommandBarButton** object. Read/write.


> [!NOTE]
> The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, see [Overview of the Office Fluent ribbon](../library-reference/concepts/overview-of-the-office-fluent-ribbon.md).


## Syntax

_expression_.**Picture**

_expression_ A variable that represents a **[CommandBarButton](Office.CommandBarButton.md)** object.


## Remarks

When you change the image on a button, you will also want to use the **Mask** property to set a mask image. The mask image determines which parts of the button image are transparent. Always set the mask after you have set the picture for a **CommandBarButton** object.

> [!NOTE]
> The images for the **View Microsoft** _Application_ and **Insert** _Item_ buttons on the **Standard** toolbar in the Visual Basic Editor cannot be changed.


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
        'Get the button image and mask of this CommandBarButton object. 
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