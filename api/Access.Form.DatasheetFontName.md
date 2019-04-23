---
title: Form.DatasheetFontName property (Access)
keywords: vbaac10.chm13396
f1_keywords:
- vbaac10.chm13396
ms.prod: access
api_name:
- Access.Form.DatasheetFontName
ms.assetid: e6b963ca-7162-912e-e63d-1437904ec8f1
ms.date: 03/12/2019
localization_priority: Normal
---


# Form.DatasheetFontName property (Access)

You can use the **DatasheetFontName** property to specify the font used to display and print field names and data in Datasheet view. Read/write **String**.


## Syntax

_expression_.**DatasheetFontName**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

The **DatasheetFontName** property applies to all fields in Datasheet view and to form controls when the form is in Datasheet view.

This property is only available in [Visual Basic](../access/Concepts/Settings/set-properties-by-using-visual-basic.md) within a Microsoft Access database.

For the **DatasheetFontName** property, the font names that you can specify depend on the fonts installed on your system and for your printer. If you specify a font that your system can't display or that isn't installed, Windows will substitute a similar font.

The following table contains the properties that don't exist in the DAO **Properties** collection until you set them by using the **Formatting (Datasheet)** toolbar, or you can add them in an Access database by using the **CreateProperty** method and append it to the **DAO Properties** collection.

|||
|:-----|:-----|
|**[DatasheetFontItalic](Access.Form.DatasheetFontItalic.md)** *|**[DatasheetForeColor](Access.Form.DatasheetForeColor.md)** *|
|**[DatasheetFontHeight](Access.Form.DatasheetFontHeight.md)** *|**[DatasheetBackColor](Access.Form.DatasheetBackColor.md)**|
|**DatasheetFontName** *|**[DatasheetGridlinesColor](Access.Form.DatasheetGridlinesColor.md)**|
|**[DatasheetFontUnderline](Access.Form.DatasheetFontUnderline.md)** *|**[DatasheetGridlinesBehavior](Access.Form.DatasheetGridlinesBehavior.md)**|
|**[DatasheetFontWeight](Access.Form.DatasheetFontWeight.md)** *|**[DatasheetCellsEffect](Access.Form.DatasheetCellsEffect.md)**|

> [!NOTE] 
> When you add or set any property listed with an asterisk, Microsoft Access automatically adds all the properties listed with an asterisk to the **Properties** collection of the database.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]