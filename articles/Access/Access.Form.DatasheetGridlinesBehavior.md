---
title: Form.DatasheetGridlinesBehavior Property (Access)
keywords: vbaac10.chm13402
f1_keywords:
- vbaac10.chm13402
ms.prod: access
api_name:
- Access.Form.DatasheetGridlinesBehavior
ms.assetid: 692268ab-69f2-4891-e460-f091b43af962
ms.date: 06/08/2017
---


# Form.DatasheetGridlinesBehavior Property (Access)

You can use the  **DatasheetGridlinesBehavior** property to specify which gridlines will appear in Datasheet view. Read/write **Byte**.


## Syntax

 _expression_. **DatasheetGridlinesBehavior**

 _expression_ A variable that represents a **Form** object.


## Remarks

This  **DatasheetGridlinesBehavior** property applies only to objects in Datasheet view.

This property is only available in [Visual Basic](set-properties-by-using-visual-basic.md)within a Microsoft Access database.

The  **DatasheetGridlinesBehavior** property uses the following settings.



|**Visual Basic**|**Description**|
|:-----|:-----|
|**acGridlinesNone**|No gridlines are displayed.|
|**acGridlinesHoriz**|Only horizontal gridlines are displayed.|
|**acGridlinesVert**|Only vertical gridlines are displayed.|
|**acGridlinesBoth**|(Default) Horizontal and vertical gridlines are displayed.|
The following table contains the properties that don't exist in the DAO  **Properties** collection of until you set them by using the **Formatting (Datasheet)** toolbar or you can add them in an Access database by using the **CreateProperty** method and append it to the DAO **Properties** collection.


|||
|:-----|:-----|
|**[DatasheetFontItalic](Access.Form.DatasheetFontItalic.md)** *|**[DatasheetForeColor](Access.Form.DatasheetForeColor.md)** *|
|**[DatasheetFontHeight](Access.Form.DatasheetFontHeight.md)** *|**[DatasheetBackColor](Access.Form.DatasheetBackColor.md)**|
|**[DatasheetFontName](Access.Form.DatasheetFontName.md)** *|**[DatasheetGridlinesColor](Access.Form.DatasheetGridlinesColor.md)**|
|**[DatasheetFontUnderline](Access.Form.DatasheetFontUnderline.md)** *|**DatasheetGridlinesBehavior**|
|**[DatasheetFontWeight](Access.Form.DatasheetFontWeight.md)** *|**[DatasheetCellsEffect](Access.Form.DatasheetCellsEffect.md)**|

 **Note**  When you add or set any property listed with an asterisk, Microsoft Access automatically adds all the properties listed with an asterisk to the  **Properties** collection in the database.


## See also


#### Concepts


[Form Object](Access.Form.md)

