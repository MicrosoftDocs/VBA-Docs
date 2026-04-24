---
title: Command (Visual C++ syntax index with import)
ms.assetid: 9c3763f1-6242-a69c-bc2a-9d885f2b122a
ms.date: 10/12/2018
ms.localizationpriority: medium
---


# Command (Visual C++ syntax index with #import)

**Applies to:** Access 2013 | Access 2016

## Methods

```csharp
HRESULT Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthadocancel_HV10294125.xml( ); 
 
_RecordsetPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcmdexecute_HV10294344.xml( VARIANT * RecordsAffected , VARIANT 
 * Parameters , long Options ); 
 
_ParameterPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdmthcreateparam_HV10294243.xml( _bstr_t Name , enum 
 DataTypeEnum Type , enum ParameterDirectionEnum Direction , long Size , 
 const _variant_t & Value  =vtMissing); 

```

## Properties

```cs
_ConnectionPtr GetActiveConnection( ); 
void PutRefActiveConnection( struct _Connection * ppvObject ); 
void PutActiveConnection( const _variant_t & ppvObject ); 
__declspec(property(get=GetActiveConnection,put=PutRefActiveConnection)) 
 _ConnectionPtr Invalid DDUE based on source, error:link not allowed in code, link filename:mdproactivecon_HV10293988.xml; 
 
_bstr_t GetCommandText( ); 
void PutCommandText( _bstr_t pbstr ); 
__declspec(property(get=GetCommandText,put=PutCommandText)) _bstr_t 
 Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocommandtext_HV10294195.xml; 
 
long GetCommandTimeout( ); 
void PutCommandTimeout( long pl ); 
__declspec(property(get=GetCommandTimeout,put=PutCommandTimeout)) long 
 Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocommandtimeout_HV10294196.xml; 
 
void PutCommandType( enum CommandTypeEnum plCmdType ); 
enum CommandTypeEnum GetCommandType( ); 
__declspec(property(get=GetCommandType,put=PutCommandType)) enum 
 CommandTypeEnum Invalid DDUE based on source, error:link not allowed in code, link filename:mdprocommandtype_HV10294197.xml; 
 
VARIANT_BOOL GetPrepared( ); 
void PutPrepared( VARIANT_BOOL pfPrepared ); 
__declspec(property(get=GetPrepared,put=PutPrepared)) VARIANT_BOOL 
 Invalid DDUE based on source, error:link not allowed in code, link filename:mdproprepared_HV10294617.xml; 
 
ParametersPtr GetParameters( ); 
__declspec(property(get=GetParameters)) ParametersPtr 
 Invalid DDUE based on source, error:link not allowed in code, link filename:mdcolparameters_HV10294594.xml; 
 
_bstr_t GetName( ); 
void PutName( _bstr_t pbstrName ); 
__declspec(property(get=GetName,put=PutName)) _bstr_t Invalid DDUE based on source, error:link not allowed in code, link filename:mdproname_HV10294535.xml; 
 
long GetState( ); 
__declspec(property(get=GetState)) long Invalid DDUE based on source, error:link not allowed in code, link filename:mdprostate_HV10294804.xml; 

```

## See also

- [Access on Microsoft Tech Community](https://techcommunity.microsoft.com/category/microsoft365/discussions/access)
- [Access Feedback Forum](https://feedbackportal.microsoft.com/feedback/forum/818e3b49-e61b-ec11-b6e7-0022481f8472)
- [Access Development on Microsoft Q&A](/answers/tags/322/m365-office-office-access-development-routing)
- [AccessForums.net](https://www.accessforums.net/index.php)
- [Access developer and VBA programming help center (FMS)](https://www.fmsinc.com/MicrosoftAccess/developer/)
- [Access posts on StackOverflow](https://stackoverflow.com/questions/tagged/ms-access)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]