---
title: Language-specific Properties, Methods, and Functions
keywords: vbaxl10.chm5278881
f1_keywords:
- vbaxl10.chm5278881
ms.prod: excel
ms.assetid: abf2101c-93ee-352b-6a67-478b9eb09003
ms.date: 06/08/2017
localization_priority: Normal
---


# Language-specific Properties, Methods, and Functions

The Excel Visual Basic for Applications (VBA) object model has language-specific elements for use with Asian and right-to-left languages.

The following table lists methods that have language-specific arguments. Methods that have new arguments or fewer arguments than in earlier versions of Excel are noted.


|**Method**|**Objects**|**Comments**|
|:-----|:-----|:-----|
| **[Add](../../api/Excel.Phonetics.Add.md)**| **Phonetics**||
| **[AddLabel](../../api/Excel.Shapes.AddLabel.md)**| **Shapes**||
| **[AddTextbox](../../api/Excel.Shapes.AddTextbox.md)**| **Shapes**||
| **AutoFormat**| **Range**||
| **[CheckSpelling](../../api/Excel.Application.CheckSpelling.md)**| **Application**,  **Chart**,  **Range**,  **Worksheet**|Added  **_SpellLang_** and removed **_IgnoreInitialAlefHamza_** and **_IgnoreFinalYaa_**|
| **[Find](../../api/Excel.Range.Find.md)**| **Application**,  **Range**|Removed  **_MatchControlCharacters_**,  **_MatchDiacritics_**,  **_MatchKashida_**, and  **_MatchAlefHamza_**|
| **[GetPhonetic](../../api/Excel.Application.GetPhonetic.md)**| **Application**||
| **[Replace](../../api/Excel.Range.Replace.md)**| **Range**|Removed  **_MatchControlCharacters_**,  **_MatchDiacritics_**,  **_MatchKashida_**, and  **_MatchAlefHamza_**|
| **[SetPhonetic](../../api/Excel.Range.SetPhonetic.md)**| **Range**||
| **[Sort](../../api/Excel.Range.Sort.md)**| **Range**|Removed  **_IgnoreControlCharacters_**,  **_IgnoreDiacritics_**, and  **_IgnoreKashida_**|
| **[SortSpecial](../../api/Excel.Range.SortSpecial.md)**| **Range**||

Properties that return or set language-specific attributes are listed in the following table.


|**Property**|**Objects**|
|:-----|:-----|
| **[AddIndent](../../api/Excel.Range.AddIndent.md)**| **Range**,  **Style**|
| **[AddressLocal](../../api/Excel.Range.AddressLocal.md)**| **Range**|
| **[Alignment](../../api/Excel.Phonetic.Alignment.md)**| **Phonetic**,  **Phonetics**,  **TextEffectFormat**,  **TickLabels**|
| **[CharacterType](../../api/Excel.Phonetic.CharacterType.md)**| **Phonetic**,  **Phonetics**|
| **[ControlCharacters](../../api/Excel.Application.ControlCharacters.md)**| **Application**|
| **[CursorMovement](../../api/Excel.Application.CursorMovement.md)**| **Application**|
| **[DefaultSheetDirection](../../api/Excel.Application.DefaultSheetDirection.md)**| **Application**|
| **[DisplayRightToLeft](../../api/Excel.Worksheet.DisplayRightToLeft.md)**| **Window**,  **Worksheet**|
| **[FileFormat](../../api/Excel.Workbook.FileFormat.md)**| **Workbook**|
| **[HorizontalAlignment](../../api/Excel.AxisTitle.HorizontalAlignment.md)**| **AxisTitle**,  **ChartTitle**,  **DataLabel**,  **DataLabels**,  **DisplayUnitLabel**,  **Range**,  **Style**,  **TextFrame**|
| **[IMEMode](../../api/Excel.Validation.IMEMode.md)**| **Validation**|
| **[International](../../api/Excel.Application.International.md)**| **Application**|
| **[Item](../../api/Excel.Phonetics.Item.md)**| **Phonetics**|
| **[Length](../../api/Excel.Phonetics.Length.md)**| **Phonetics**|
| **[Phonetic](../../api/Excel.Range.Phonetic.md)**| **Range**|
| **[PhoneticCharacters](../../api/Excel.Characters.PhoneticCharacters.md)**| **Characters**|
| **[Phonetics](../../api/Excel.Range.Phonetics.md)**| **Range**|
| **[ReadingOrder](../../api/Excel.AxisTitle.ReadingOrder.md)**| **AxisTitle**,  **ChartTitle**,  **DataLabel**,  **DataLabels**,  **DisplayUnitLabel**,  **Range**,  **Style**,  **TextFrame**,  **TickLabels**|
| **[Start](../../api/Excel.Phonetics.Start.md)**| **Phonetics**|
| **[VerticalAlignment](../../api/Excel.AxisTitle.VerticalAlignment.md)**| **AxisTitle**,  **ChartTitle**,  **DataLabel**,  **DataLabels**,  **DisplayUnitLabels**,  **Range**,  **Style**,  **TextFrame**|

The following are language-specific worksheet functions:

-  **[FindB](./Events-WorksheetFunctions-Shapes/list-of-worksheet-functions-available-to-visual-basic.md)**
    
-  **[ReplaceB](./Events-WorksheetFunctions-Shapes/list-of-worksheet-functions-available-to-visual-basic.md)**
    
-  **[SearchB](./Events-WorksheetFunctions-Shapes/list-of-worksheet-functions-available-to-visual-basic.md)**
    
-  **[USDollar](./Events-WorksheetFunctions-Shapes/list-of-worksheet-functions-available-to-visual-basic.md)**

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
