---
title: Image object (Outlook Forms Script)
keywords: olfm10.chm2000540
f1_keywords:
- olfm10.chm2000540
ms.prod: outlook
ms.assetid: d2bcc281-6af0-5bbf-fa7f-ac581dbcf5dc
ms.date: 06/08/2017
localization_priority: Normal
---


# Image object (Outlook Forms Script)

Displays a picture on a form.


## Remarks

The **Image** control lets you display a picture as part of the data in a form. For example, you might use an **Image** to display employee photographs in a personnel form.

The **Image** lets you crop, size, or zoom a picture, but does not allow you to edit the contents of the picture. For example, you cannot use the **Image** to change the colors in the picture, to make the picture transparent, or to refine the image of the picture. You must use image editing software for these purposes.

The **Image** object supports the following file formats:

- *.bmp   
- *.cur    
- *.gif   
- *.ico    
- *.jpg   
- *.wmf
    

You can also display a picture on a **[Label](Outlook.label.md)**. However, a **Label** does not let you crop, size, or zoom the picture.


## Events

|Name|Description|
|:-----|:-----|
| [Click](Outlook.Image.click.md)|Occurs when the user clicks inside the control.|


## Properties

|Name|Description|
|:-----|:-----|
| [AutoSize](Outlook.Image.md)|Returns or sets a **Boolean** that specifies whether an object automatically resizes to display its entire contents. Read/write.|
| [BackColor](Outlook.Image.backcolor.md)|Returns or sets a **Long** that specifies the background color of the object. Read/write.|
| [BackStyle](Outlook.Image.backstyle.md)|Returns or sets an **Integer** that specifies the background style for an object. Read/write.|
| [BorderColor](Outlook.Image.bordercolor.md)|Returns or sets a **Long** that specifies the border color of an object. Read/write.|
| [BorderStyle](Outlook.Image.borderstyle.md)|Returns or sets an **Integer** that specifies the type of border of the control. Read/write.|
| [Enabled](Outlook.Image.enabled.md)|Returns or sets a **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.|
| [MouseIcon](Outlook.Image.mouseicon.md)|Returns a **String** that represents the full path name of a custom icon that is to be assigned to the control. Read-only.|
| [MousePointer](Outlook.Image.mousepointer.md)|Returns or sets an **Integer** that specifies the type of pointer displayed when the user positions the mouse over a particular object. Read/write.|
| [Picture](Outlook.Image.picture.md)|Returns a **String** that specifies the full path name of a bitmap to display on a control. Read-only.|
| [PictureAlignment](Outlook.Image.picturealignment.md)|Returns or sets an **Integer** that specifies the location of a background picture. Read/write.|
| [PictureSizeMode](Outlook.Image.picturesizemode.md)|Returns or sets an **Integer** that specifies how to display the background picture on a control. Read/write.|
| [PictureTiling](Outlook.Image.picturetiling.md)|Returns or sets a **Boolean** that specifies whether a picture is repeated across the background of the object. Read/write.|
| [SpecialEffect](Outlook.Image.specialeffect.md)|Returns or sets an **Integer** that specifies the visual appearance of an object. Read/write.|





[!include[Support and feedback](~/includes/feedback-boilerplate.md)]