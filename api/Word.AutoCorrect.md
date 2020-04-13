---
title: AutoCorrect object (Word)
keywords: vbawd10.chm2377
f1_keywords:
- vbawd10.chm2377
ms.prod: word
api_name:
- Word.AutoCorrect
ms.assetid: dea9b72c-4378-05ac-ec4b-51cf3af3f2a3
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCorrect object (Word)

Represents the AutoCorrect functionality in Word.


## Remarks

Use the **[AutoCorrect](Word.Application.AutoCorrect.md)** property to return the **AutoCorrect** object. The following example enables the AutoCorrect options and creates an AutoCorrect entry.


```vb
With AutoCorrect 
 .CorrectCapsLock = True 
 .CorrectDays = True 
 .Entries.Add Name:="usualy", Value:="usually" 
End With
```

The **[Entries](Word.AutoCorrect.Entries.md)** property returns the **[Entries](Word.AutoCorrect.Entries.md)** object that represents the AutoCorrect entries in the **AutoCorrect** dialog box.

## Properties

- [Application](Word.AutoCorrect.Application.md)
- [CorrectCapsLock](Word.AutoCorrect.CorrectCapsLock.md)
- [CorrectDays](Word.AutoCorrect.CorrectDays.md)
- [CorrectHangulAndAlphabet](Word.AutoCorrect.CorrectHangulAndAlphabet.md)
- [CorrectInitialCaps](Word.AutoCorrect.CorrectInitialCaps.md)
- [CorrectKeyboardSetting](Word.AutoCorrect.CorrectKeyboardSetting.md)
- [CorrectSentenceCaps](Word.AutoCorrect.CorrectSentenceCaps.md)
- [CorrectTableCells](Word.AutoCorrect.CorrectTableCells.md)
- [Creator](Word.AutoCorrect.Creator.md)
- [DisplayAutoCorrectOptions](Word.AutoCorrect.DisplayAutoCorrectOptions.md)
- [Entries](Word.AutoCorrect.Entries.md)
- [FirstLetterAutoAdd](Word.AutoCorrect.FirstLetterAutoAdd.md)
- [FirstLetterExceptions](Word.AutoCorrect.FirstLetterExceptions.md)
- [HangulAndAlphabetAutoAdd](Word.AutoCorrect.HangulAndAlphabetAutoAdd.md)
- [HangulAndAlphabetExceptions](Word.AutoCorrect.HangulAndAlphabetExceptions.md)
- [OtherCorrectionsAutoAdd](Word.AutoCorrect.OtherCorrectionsAutoAdd.md)
- [OtherCorrectionsExceptions](Word.AutoCorrect.OtherCorrectionsExceptions.md)
- [Parent](Word.AutoCorrect.Parent.md)
- [ReplaceText](Word.AutoCorrect.ReplaceText.md)
- [ReplaceTextFromSpellingChecker](Word.AutoCorrect.ReplaceTextFromSpellingChecker.md)
- [TwoInitialCapsAutoAdd](Word.AutoCorrect.TwoInitialCapsAutoAdd.md)
- [TwoInitialCapsExceptions](Word.AutoCorrect.TwoInitialCapsExceptions.md)

## See also

- [Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]