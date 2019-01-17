---
title: DefaultWebOptions.FolderSuffix property (Excel)
keywords: vbaxl10.chm660089
f1_keywords:
- vbaxl10.chm660089
ms.prod: excel
api_name:
- Excel.DefaultWebOptions.FolderSuffix
ms.assetid: ff0821ab-a2fd-58bc-058c-2abdaefbf04d
ms.date: 06/08/2017
localization_priority: Normal
---


# DefaultWebOptions.FolderSuffix property (Excel)

Returns the folder suffix that Microsoft Excel uses when you save a document as a Web page, use long file names, and choose to save supporting files in a separate folder (that is, if the  **[UseLongFileNames](Excel.DefaultWebOptions.UseLongFileNames.md)** and **[OrganizeInFolder](Excel.DefaultWebOptions.OrganizeInFolder.md)** properties are set to **True**). Read-only **String**.


## Syntax

_expression_. `FolderSuffix`

_expression_ A variable that represents a [DefaultWebOptions](Excel.DefaultWebOptions.md) object.


## Remarks

Newly created documents use the suffix returned by the  **FolderSuffix** property of the **DefaultWebOptions** object. The value of the **FolderSuffix** property of the **WebOptions** object may differ from that of the **DefaultWebOptions** object if the document was previously edited in a different language version of Microsoft Excel. You can use the **[UseDefaultFolderSuffix](Excel.WebOptions.UseDefaultFolderSuffix.md)** method to change the suffix to the language you are currently using in Microsoft Office.

By default, the name of the supporting folder is the name of the Web page plus an underscore (_), a period (.), or a hyphen (-) and the word "files" (appearing in the language of the version of Excel in which the file was saved as a Web page). For example, suppose that you use the Dutch language version of Excel to save a file called "Page1" as a Web page. The default name of the supporting folder is Page1_bestanden.

The following table lists each language version of Office, and gives its corresponding  **LanguageID** property value and folder suffix. For the languages that are not listed in the table, the suffix ".files" is used.



|**Language**|**LanguageID**|**Folder suffix**|
|:-----|:-----|:-----|
|Arabic|1025|.files|
|Basque (Basque)|1069|_fitxategiak|
|Portuguese (Brazil)|1046|_arquivos|
|Bulgarian|1026|.files|
|Catalan|1027|_fitxers|
|Chinese - Simplified|2052|.files|
|Chinese - Traditional|1028|.files|
|Croatian|1050|_datoteke|
|Czech|1029|_soubory|
|Danish|1030|-filer|
|Dutch|1043|_bestanden|
|English|1033|_files|
|Estonian|1061|_failid|
|Finnish|1035|_tiedostot|
|French|1036|_fichiers|
|German|1031|-Dateien|
|Greek|1032|.files|
|Hebrew|1037|.files|
|Hungarian|1038|_elemei|
|Italian|1040|-file|
|Japanese|1041|.files|
|Korean|1042|.files|
|Latvian|1062|_fails|
|Lithuanian|1063|_bylos|
|Norwegian|1044|-filer|
|Polish|1045|_pliki|
|Portuguese|2070|_ficheiros|
|Romanian|1048|.files|
|Russian|1049|.files|
|Serbian (Cyrillic)|3098|.files|
|Serbian (Latin)|2074|_fajlovi|
|Slovakian|1051|.files|
|Slovenian|1060|_datoteke|
|Spanish|3082|_archivos|
|Swedish|1053|-filer|
|Thai|1054|.files|
|Turkish|1055|_dosyalar|
|Ukranian|1058|.files|
|Vietnamese|1066|.files|

## See also


[DefaultWebOptions Object](Excel.DefaultWebOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]