---
title: Dictionary object (Word)
keywords: vbawd10.chm2477
f1_keywords:
- vbawd10.chm2477
ms.prod: word
api_name:
- Word.Dictionary
ms.assetid: 1946d60c-2abd-9ca9-8d0b-7068e9173bb3
ms.date: 06/08/2017
localization_priority: Normal
---


# Dictionary object (Word)

Represents a dictionary.  **Dictionary** objects that represent custom dictionaries are members of the **[Dictionaries](Word.dictionaries.md)** collection. Other dictionary objects are returned by properties of the **[Languages](Word.languages.md)** collection; these include the **[ActiveSpellingDictionary](Word.Language.ActiveSpellingDictionary.md)**, **[ActiveGrammarDictionary](Word.Language.ActiveGrammarDictionary.md)**, **[ActiveThesaurusDictionary](Word.Language.ActiveThesaurusDictionary.md)**, and **[ActiveHyphenationDictionary](Word.Language.ActiveHyphenationDictionary.md)** properties.


## Remarks

Use  **[CustomDictionaries](Word.Application.CustomDictionaries.md)** (Index), where Index is an index number or the string name for the dictionary, to return a single **Dictionary** object that represents a custom dictionary. The following example returns the first dictionary in the collection.


```vb
CustomDictionaries(1)
```

The following example returns the dictionary named "MyDictionary."




```vb
CustomDictionaries("MyDictionary")
```

Use the **[ActiveCustomDictionary](Word.Dictionaries.ActiveCustomDictionary.md)** property to set the custom spelling dictionary in the collection to which new words are added. If you try to set this property to a dictionary that's not a custom dictionary, an error occurs.

Use the **[Add](Word.Dictionaries.Add.md)** method to add a new dictionary to the collection of active custom dictionaries. If there is no file with the name specified by FileName, Word creates it. The following example adds "MyCustom.dic" to the collection of custom dictionaries.




```vb
CustomDictionaries.Add FileName:="MyCustom.dic"
```

Remarks

Use the **[Name](Word.Dictionary.Name.md)** and **[Path](Word.Dictionary.Path.md)** properties to locate any of the dictionaries. The following example displays a message box that contains the full path for each dictionary.




```vb
For Each d in CustomDictionaries 
 Msgbox d.Path & Application.PathSeparator & d.Name 
Next d
```

Use the **[LanguageSpecific](Word.Dictionary.LanguageSpecific.md)** property to determine whether the specified custom dictionary can have a specific language assigned to it with the **[LanguageID](Word.Dictionary.LanguageID.md)** property. If the dictionary is language specific, it will verify only text that's formatted for the specified language.

For each language for which proofing tools are installed, you can use the **ActiveGrammarDictionary**, **ActiveHyphenationDictionary**, **ActiveSpellingDictionary**, and **ActiveThesaurusDictionary** properties to return the corresponding **Dictionary** objects. The following example returns the full path for the active spelling dictionary used in the U.S. English version of Word.




```vb
Set myspell = Languages(wdEnglishUS).ActiveSpellingDictionary 
MsgBox mySpell.Path & Application.PathSeparator & mySpell.Name
```

The **[ReadOnly](Word.Dictionary.ReadOnly.md)** property returns **True** for .lex files (built-in proofing dictionaries) and **False** for .dic files (custom spelling dictionaries).


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]