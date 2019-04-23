---
title: IMEStatus function (Visual Basic for Applications)
keywords: vblr6.chm1011064
f1_keywords:
- vblr6.chm1011064
ms.prod: office
ms.assetid: 8fb525b1-0243-79c8-32fc-4eb8d634e351
ms.date: 12/13/2018
localization_priority: Normal
---


# IMEStatus function

Returns an [Integer](../../Glossary/vbe-glossary.md#integer-data-type) specifying the current Input Method Editor (IME) mode of Microsoft Windows; available in East Asian versions only.

## Syntax

**IMEStatus**

## Return values

The return values for the Japanese [locale](../../Glossary/vbe-glossary.md#locale) are as follows.

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbIMEModeNoControl**|0|Don't control IME (default)|
|**vbIMEModeOn**|1|IME on|
|**vbIMEModeOff**|2|IME off|
|**vbIMEModeDisable**|3|IME disabled|
|**vbIMEModeHiragana**|4|Full-width Hiragana mode|
|**vbIMEModeKatakana**|5|Full-width Katakana mode|
|**vbIMEModeKatakanaHalf**|6|Half-width Katakana mode|
|**vbIMEModeAlphaFull**|7|Full-width Alphanumeric mode|
|**vbIMEModeAlpha**|8|Half-width Alphanumeric mode|

<br/>

The return values for the Korean locale are as follows.

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbIMEModeNoControl**|0|Don't control IME(default)|
|**vbIMEModeAlphaFull**|7|Full-width Alphanumeric mode|
|**vbIMEModeAlpha**|8|Half-width Alphanumeric mode|
|**vbIMEModeHangulFull**|9|Full-width Hangul mode|
|**vbIMEModeHangul**|10|Half-width Hangul mode|

<br/>

The return values for the Chinese locale are as follows.

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbIMEModeNoControl**|0|Don't control IME (default)|
|**vbIMEModeOn**|1|IME on|
|**vbIMEModeOff**|2|IME off|

## See also

- [Functions (Visual Basic for Applications)](../functions-visual-basic-for-applications.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]