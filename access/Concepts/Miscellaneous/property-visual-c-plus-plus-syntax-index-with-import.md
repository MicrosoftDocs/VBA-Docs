---
title: Property (Visual C++ syntax index with import)
ms.assetid: 3649db1b-ab2f-2767-a8b3-a146720217c0
ms.date: 10/12/2018
ms.localizationpriority: medium
---


# Property (Visual C++ syntax index with #import)

**Applies to:** Access 2013 | Access 2016

## Properties

```cs
long GetAttributes( ); 
void PutAttributes( long plAttributes ); 
__declspec(property(get=GetAttributes,put=PutAttributes)) long 
 Invalid DDUE based on source, error:link not allowed in code, link filename:adproattributes_HV10294099.xml; 
 
_bstr_t GetName( ); 
__declspec(property(get=GetName)) _bstr_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdproname_HV10294535.xml; 
 
enum DataTypeEnum GetType( ); 
__declspec(property(get=GetType)) enum DataTypeEnum Invalid DDUE based on source, error:link not allowed in code, link filename:mdprotype_HV10294866.xml; 
 
_variant_t GetValue( ); 
void PutValue( const _variant_t & pval ); 
__declspec(property(get=GetValue,put=PutValue)) _variant_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdprovalue_HV10294920.xml; 

```

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]