---
title: Form.DatasheetCellsEffect property (Access)
keywords: vbaac10.chm13404
f1_keywords:
- vbaac10.chm13404
ms.prod: access
api_name:
- Access.Form.DatasheetCellsEffect
ms.assetid: 3820b218-37b0-d5b5-bae2-8a179cc9b87a
ms.date: 03/12/2019
localization_priority: Normal
---


# Form.DatasheetCellsEffect property (Access)

You can use the **DatasheetCellsEffect** property to specify whether special effects are applied to cells in a datasheet. Read/write **Byte**.


## Syntax

_expression_.**DatasheetCellsEffect**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

The **DatasheetCellsEffect** property applies only to objects in Datasheet view.

This property is only available within a Microsoft Access database.

The **DatasheetCellsEffect** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|Flat|**acEffectNormal**|(Default) No special effects are applied to the cells in the datasheet.|
|Raised|**acEffectRaised**|Cells in the datasheet appear raised.|
|Sunken|**acEffectSunken**|Cells in the datasheet appear sunken.|

This property applies the selected effect to the entire datasheet.

When this property is set to Raised or Sunken, gridlines will be visible on the datasheet regardless of the **[DatasheetGridlinesBehavior](Access.Form.DatasheetGridlinesBehavior.md)** property setting.

The following table contains the properties that don't exist in the DAO **Properties** collection until you set them by using the **Formatting (Datasheet)** toolbar, or you can add them in an Access database (.mdb) by using the **CreateProperty** method and append it to the DAO **Properties** collection.

|||
|:-----|:-----|
|**[DatasheetFontItalic](Access.Form.DatasheetFontItalic.md)** *|**[DatasheetForeColor](Access.Form.DatasheetForeColor.md)** *|
|**[DatasheetFontHeight](Access.Form.DatasheetFontHeight.md)** *|**[DatasheetBackColor](Access.Form.DatasheetBackColor.md)**|
|**[DatasheetFontName](Access.Form.DatasheetFontName.md)** *|**[DatasheetGridlinesColor](Access.Form.DatasheetGridlinesColor.md)**|
|**[DatasheetFontUnderline](Access.Form.DatasheetFontUnderline.md)** *|**[DatasheetGridlinesBehavior](Access.Form.DatasheetGridlinesBehavior.md)**|
|**[DatasheetFontWeight](Access.Form.DatasheetFontWeight.md)** *|**DatasheetCellsEffect**|

> [!NOTE] 
> When you add or set any property listed with an asterisk, Microsoft Access automatically adds all the properties listed with an asterisk to the **Properties** collection in the database.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]