unit V8AddIn;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, CTypes, Variants, LCLProc, contnrs;

type

  TAppCapabilities = (
    eAppCapabilitiesInvalid = -1,
    eAppCapabilities1 = 1);

  TAppType = (
    eAppUnknown = -1,
    eAppThinClient = 0,
    eAppThickClient,
    eAppWebClient,
    eAppServer,
    eAppExtConn);

  TErrorCode = (
    ADDIN_E_NONE = 1000,
    ADDIN_E_ORDINARY = 1001,
    ADDIN_E_ATTENTION = 1002,
    ADDIN_E_IMPORTANT = 1003,
    ADDIN_E_VERY_IMPORTANT = 1004,
    ADDIN_E_INFO = 1005,
    ADDIN_E_FAIL = 1006,
    ADDIN_E_MSGBOX_ATTENTION = 1007,
    ADDIN_E_MSGBOX_INFO = 1008,
    ADDIN_E_MSGBOX_FAIL = 1009);

  { TAddIn }

  TAddIn = class (Tobject)
  public type
    TGetMethod = function: Variant of object;
    TSetMethod = procedure (AValue: Variant) of object;
    TProcMethod = procedure (const Params: array of Variant) of object;
    TProc0Method = procedure of object;
    TProc1Method = procedure (P1: Variant) of object;
    TProc2Method = procedure (P1, P2: Variant) of object;
    TProc3Method = procedure (P1, P2, P3: Variant) of object;
    TProc4Method = procedure (P1, P2, P3, P4: Variant) of object;
    TProc5Method = procedure (P1, P2, P3, P4, P5: Variant) of object;
    TFuncMethod = function (const Params: array of Variant): Variant of object;
    TFunc0Method = function: Variant of object;
    TFunc1Method = function (P1: Variant): Variant of object;
    TFunc2Method = function (P1, P2: Variant): Variant of object;
    TFunc3Method = function (P1, P2, P3: Variant): Variant of object;
    TFunc4Method = function (P1, P2, P3, P4: Variant): Variant of object;
    TFunc5Method = function (P1, P2, P3, P4, P5: Variant): Variant of object;
  private
    FConnection: Pointer;
    FMemoryManager: Pointer;
    FFactory: Pointer;
  protected
    class function AppCapabilities: TAppCapabilities; static;
    { IMemoryManager }
    function AllocMemory(Size: SizeInt): Pointer;
    procedure FreeMemory(var P: Pointer);
    { IAddInDefBase }
    function AddError(Code: TErrorCode; const Source, Descr: String; HRes: HRESULT): Boolean;
    function Read(const PropName: String; var Value: Variant; var ErrCode: LongWord; var ErrDescriptor: String): Boolean;
    function Write(const PropName: PWideChar; const Value: Variant): Boolean;
    function RegisterProfileAs(const ProfileName: String): Boolean;
    function SetEventBufferDepth(const Depth: Integer): Boolean;
    function GetEventBufferDepth: Integer;
    function ExternalEvent(const Source, Message, Data: String): Boolean;
    procedure CleanEventBuffer;
    function SetStatusLine(const StatusLine: String): Boolean;
    procedure ResetStatusLine;
    { IPlatformInfo }
    function AppVersion: String;
    function UserAgentInformation: String;
    function ApplicationType: TAppType;
    { IMsgBox }
    function Confirm(const QueryText: String; var RetValue: Variant): Boolean; overload;
    function Confirm(const QueryText: String): Boolean; overload;
    function Alert(const Text: String): Boolean;
  public
    constructor Create; virtual;
    { IInitDoneBase }
    function Init: Boolean; virtual;
    function GetInfo: Integer; virtual;
    procedure Done; virtual;
    { ILanguageExtenderBase }
    function RegisterExtensionAs(var ExtensionName: String): Boolean; virtual;
    function GetNProps: Integer; virtual;
    function FindProp(const PropName: String): Integer; virtual;
    function GetPropName(const PropNum, PropAlias: Integer): String; virtual;
    function GetPropVal(const PropNum: Integer; var Value: Variant): Boolean; virtual;
    function SetPropVal(const PropNum: Integer; const Value: Variant): Boolean; virtual;
    function IsPropReadable(const PropNum: Integer): Boolean; virtual;
    function IsPropWritable(const PropNum: Integer): Boolean; virtual;
    function GetNMethods: Integer; virtual;
    function FindMethod(const AMethodName: String): Integer; virtual;
    function GetMethodName(const MethodNum, MethodAlias: Integer): String; virtual;
    function GetNParams(const MethodNum: Integer): Integer; virtual;
    function GetParamDefValue(const MethodNum, ParamNum: Integer; var Value: Variant): Boolean; virtual;
    function HasRetVal(const MethodNum: Integer): Boolean; virtual;
    function CallAsProc(const MethodNum: Integer; const Params: array of Variant): Boolean; virtual;
    function CallAsFunc(const MethodNum: Integer; var RetValue: Variant; const Params: array of Variant): Boolean; virtual;
    { ILocaleBase }
    procedure SetLocale(const Locale: String); virtual;
    { Easy use }
    class procedure RegisterAddInClass(const AddInExtensionName: String);
    class procedure AddProp(const Name, NameAlias: String; const ReadMethod: TGetMethod = nil; const WriteMethod: TSetMethod = nil);
    class procedure AddProc(const Name, NameAlias: String; const Method: TProcMethod; const NParams: Integer); overload;
    class procedure AddProc(const Name, NameAlias: String; const Method: TProc0Method); overload;
    class procedure AddProc(const Name, NameAlias: String; const Method: TProc1Method); overload;
    class procedure AddProc(const Name, NameAlias: String; const Method: TProc2Method); overload;
    class procedure AddProc(const Name, NameAlias: String; const Method: TProc3Method); overload;
    class procedure AddProc(const Name, NameAlias: String; const Method: TProc4Method); overload;
    class procedure AddProc(const Name, NameAlias: String; const Method: TProc5Method); overload;
    class procedure AddFunc(const Name, NameAlias: String; const Method: TFuncMethod; const NParams: Integer); overload;
    class procedure AddFunc(const Name, NameAlias: String; const Method: TFunc0Method); overload;
    class procedure AddFunc(const Name, NameAlias: String; const Method: TFunc1Method); overload;
    class procedure AddFunc(const Name, NameAlias: String; const Method: TFunc2Method); overload;
    class procedure AddFunc(const Name, NameAlias: String; const Method: TFunc3Method); overload;
    class procedure AddFunc(const Name, NameAlias: String; const Method: TFunc4Method); overload;
    class procedure AddFunc(const Name, NameAlias: String; const Method: TFunc5Method); overload;
  end;

  TAddInClass = class of TAddIn;

procedure Unused(const AValue);

procedure RegisterAddInClass(const AddInClass: TAddInClass);

{****************************************************************************
                        1C binary data
****************************************************************************}

function StreamTo1CBinaryData(const Stream: TStream): Variant;
function Var1CBinaryData: TVarType;
function Is1CBinaryData(const Value: Variant): Boolean;
function Get1CBinaryDataSize(const Value: Variant): SizeInt;
function Get1CBinaryData(const Value: Variant): Pointer;

{****************************************************************************
                        Exported functions
****************************************************************************}

function GetClassNames: PWideChar; cdecl;
function GetClassObject(const wsName: PWideChar; var pIntf: Pointer): clong; cdecl;
function DestroyObject(var pIntf: Pointer): LongInt; cdecl;
function SetPlatformCapabilities(const Capabilities: TAppCapabilities): TAppCapabilities; cdecl;

implementation

// Use to hide "parameter not used" hints.
{$PUSH}{$HINTS OFF}
procedure Unused(const AValue);
begin
end;
{$POP}

type
  TAddInFactory = class;

{$I types.inc}
{$I componentbase.inc}
{$I addindefbase.inc}
{$I imemorymanager.inc}
{$I strings.inc}
{$I binarydata.inc}
{$I variants.inc}
{$I bindings.inc}

var
  ClassNames: PWideChar;
  ComponentBase: TComponentBase;
  AppCapabilities: TAppCapabilities;
  FactoryList: TStringList;

{$I factory.inc}

procedure RegisterAddInClass(const AddInClass: TAddInClass);
var
  Index: Integer;
  Factory: TAddInFactory;
begin
  Index := FactoryList.IndexOf(AddInClass.ClassName);
  if Index >= 0 then
    FactoryList.Delete(Index);
  Factory := TAddInFactory.Create;
  Factory.AddInClass := AddInClass;
  FactoryList.AddObject(AddInClass.ClassName, Factory);
end;

{****************************************************************************
                        Exported functions
****************************************************************************}

function GetClassNames: PWideChar; cdecl;
begin
  if ClassNames = nil then
    ClassNames := StringToWideChar(FactoryList.DelimitedText, nil);
  Result := ClassNames;
end;

function GetClassNamesS: String;
begin
  Result := FactoryList.DelimitedText;
end;

function GetClassObject(const wsName: PWideChar; var pIntf: Pointer): clong;
  cdecl;
var
  ClassName: String;
  Index: Integer;
  AddInFactory: TAddInFactory;
begin
  Result := 0;
  if pIntf <> nil then
    Exit;
  ClassName := WideCharToString(wsName);
  Index := FactoryList.IndexOf(ClassName);
  if Index < 0 then
    Exit;
  AddInFactory := TAddInFactory(FactoryList.Objects[Index]);
  pIntf := AddInFactory.New;
  {$PUSH}
  {$HINTS OFF}
  Result := PtrInt(pIntf);
  {$POP}
end;

function DestroyObject(var pIntf: Pointer): LongInt; cdecl;
begin
  Result := -1;
  if pIntf = nil then
    Exit;
  PComponentBase(pIntf)^.Factory.Delete(pIntf);
  Result := 0;
end;

function SetPlatformCapabilities(const Capabilities: TAppCapabilities
  ): TAppCapabilities; cdecl;
begin
  AppCapabilities := Capabilities;
  Result := eAppCapabilities1;
end;

{****************************************************************************
                        TAddIn class
****************************************************************************}

{ TAddIn }

{$PUSH}
{$MACRO ON}

{$define ClassFactory := TAddInFactory(FFactory)}

class function TAddIn.AppCapabilities: TAppCapabilities;
begin
  Result := V8AddIn.AppCapabilities;
end;

function TAddIn.AllocMemory(Size: SizeInt): Pointer;
begin
  Result := nil;
  if FMemoryManager = nil then
    Exit;
  if not P1CMemoryManager(FMemoryManager)^.__vfptr^.AllocMemory(FMemoryManager, Result, Size) then
    Result := nil;
end;

procedure TAddIn.FreeMemory(var P: Pointer);
begin
  P1CMemoryManager(FMemoryManager)^.__vfptr^.FreeMemory(FMemoryManager, P);
end;

function TAddIn.AddError(Code: TErrorCode; const Source, Descr: String;
  HRes: HRESULT): Boolean;
var
  S: PWideChar;
  D: PWideChar;
begin
  S := StringToWideChar(Source, nil);
  D := StringToWideChar(Descr, nil);
  Result := PAddInDefBase(FConnection)^.__vfptr^.AddError(FConnection, Ord(Code), S, D, HRes);
  FreeMem(S);
  FreeMem(D);
end;

function TAddIn.Read(const PropName: String; var Value: Variant;
  var ErrCode: LongWord; var ErrDescriptor: String): Boolean;
var
  wsPropName: PWideChar;
  Val: T1CVariant;
  E: culong;
  wsDesc: PWideChar;
begin
  wsPropName := StringToWideChar(PropName, nil);
  Result := PAddInDefBase(FConnection)^.__vfptr^.Read(FConnection, wsPropName, @Val, @E, @wsDesc);
  if Result then
    begin
      From1CVariant(Val, Value);
      ErrCode := E;
      ErrDescriptor := WideCharToString(wsDesc);
    end
  else
    begin
      Value := Unassigned;
      ErrCode := 0;
      ErrDescriptor := '';
    end;
  FreeMem(wsPropName);
end;

function TAddIn.Write(const PropName: PWideChar; const Value: Variant): Boolean;
var
  wsPropName: PWideChar;
  Val: T1CVariant;
begin
  wsPropName := StringToWideChar(PropName, nil);
  Val.vt := VTYPE_EMPTY;
  To1CVariant(Value, Val, FMemoryManager);
  Result := PAddInDefBase(FConnection)^.__vfptr^.Write(FConnection, wsPropName, @Val);
  FreeMem(wsPropName);
end;

function TAddIn.RegisterProfileAs(const ProfileName: String): Boolean;
var
  wsProfileName: PWideChar;
begin
  wsProfileName := StringToWideChar(ProfileName, nil);
  Result := PAddInDefBase(FConnection)^.__vfptr^.RegisterProfileAs(FConnection, wsProfileName);
  FreeMem(wsProfileName);
end;

function TAddIn.SetEventBufferDepth(const Depth: Integer): Boolean;
begin
  Result := PAddInDefBase(FConnection)^.__vfptr^.SetEventBufferDepth(FConnection, Depth);
end;

function TAddIn.GetEventBufferDepth: Integer;
begin
  Result := PAddInDefBase(FConnection)^.__vfptr^.GetEventBufferDepth(FConnection);
end;

function TAddIn.ExternalEvent(const Source, Message, Data: String): Boolean;
var
  wsSource, wsMessage, wsData: PWideChar;
begin
  wsSource := StringToWideChar(Source, nil);
  wsMessage := StringToWideChar(Message, nil);
  wsData := StringToWideChar(Data, nil);
  Result := PAddInDefBase(FConnection)^.__vfptr^.ExternalEvent(FConnection, wsSource, wsMessage, wsData);
  FreeMem(wsSource);
  FreeMem(wsMessage);
  FreeMem(wsData);
end;

procedure TAddIn.CleanEventBuffer;
begin
  PAddInDefBase(FConnection)^.__vfptr^.CleanEventBuffer(FConnection);
end;

function TAddIn.SetStatusLine(const StatusLine: String): Boolean;
var
  wsStatusLine: PWideChar;
begin
  wsStatusLine := StringToWideChar(StatusLine, nil);
  Result := PAddInDefBase(FConnection)^.__vfptr^.SetStatusLine(FConnection, wsStatusLine);
  FreeMem(wsStatusLine);
end;

procedure TAddIn.ResetStatusLine;
begin
  PAddInDefBase(FConnection)^.__vfptr^.ResetStatusLine(FConnection);
end;

function TAddIn.AppVersion: String;
var
  IFace: PPlatformInfo;
  AppInfo: PAppInfo;
begin
  IFace := PPlatformInfo(PAddInDefBaseEx(FConnection)^.__vfptr^.GetInterface(FConnection, eIPlatformInfo) - SizeOf(Pointer));
  AppInfo := IFace^.__vfptr^.GetPlatformInfo(IFace);
  Result := WideCharToString(AppInfo^.AppVersion);
end;

function TAddIn.UserAgentInformation: String;
var
  IFace: PPlatformInfo;
  AppInfo: PAppInfo;
begin
  IFace := PPlatformInfo(PAddInDefBaseEx(FConnection)^.__vfptr^.GetInterface(FConnection, eIPlatformInfo) - SizeOf(Pointer));
  AppInfo := IFace^.__vfptr^.GetPlatformInfo(IFace);
  Result := WideCharToString(AppInfo^.UserAgentInformation);
end;

function TAddIn.ApplicationType: TAppType;
var
  IFace: PPlatformInfo;
  AppInfo: PAppInfo;
begin
  IFace := PPlatformInfo(PAddInDefBaseEx(FConnection)^.__vfptr^.GetInterface(FConnection, eIPlatformInfo) - SizeOf(Pointer));
  AppInfo := IFace^.__vfptr^.GetPlatformInfo(IFace);
  Result := AppInfo^.Application;
end;

function TAddIn.Confirm(const QueryText: String; var RetValue: Variant
  ): Boolean;
var
  wsQueryText: PWideChar;
  RV: T1CVariant;
  IFace: PMsgBox;
begin
  RV.vt := VTYPE_EMPTY;
  RetValue := Unassigned;
  To1CVariant(RetValue, RV, nil);
  wsQueryText := StringToWideChar(QueryText, nil);
  IFace := PMsgBox(PAddInDefBaseEx(FConnection)^.__vfptr^.GetInterface(FConnection, eIMsgBox) - SizeOf(Pointer));
  Result := IFace^.__vfptr^.Confirm(IFace, wsQueryText, @RV);
  if Result then
    From1CVariant(RV, RetValue);
  FreeMem(wsQueryText);
end;

function TAddIn.Confirm(const QueryText: String): Boolean;
var
  wsQueryText: PWideChar;
  RV: T1CVariant;
  IFace: PMsgBox;
  RetValue: Variant;
begin
  RV.vt := VTYPE_EMPTY;
  RetValue := Unassigned;
  To1CVariant(RetValue, RV, nil);
  wsQueryText := StringToWideChar(QueryText, nil);
  IFace := PMsgBox(PAddInDefBaseEx(FConnection)^.__vfptr^.GetInterface(FConnection, eIMsgBox) - SizeOf(Pointer));
  Result := IFace^.__vfptr^.Confirm(IFace, wsQueryText, @RV);
  if Result then
    begin
      From1CVariant(RV, RetValue);
      Result := RetValue;
    end;
  FreeMem(wsQueryText);
end;

function TAddIn.Alert(const Text: String): Boolean;
var
  wsText: PWideChar;
  IFace: PMsgBox;
begin
  wsText := StringToWideChar(Text, nil);
  IFace := PMsgBox(PAddInDefBaseEx(FConnection)^.__vfptr^.GetInterface(FConnection, eIMsgBox) - SizeOf(Pointer));
  Result := IFace^.__vfptr^.Alert(IFace, wsText);
  FreeMem(wsText);
end;

constructor TAddIn.Create;
begin
  // Do noting by default
end;

function TAddIn.Init: Boolean;
begin
  Result := True;
end;

function TAddIn.GetInfo: Integer;
begin
  Result := 2000;
end;

procedure TAddIn.Done;
begin
  // Do noting by default
end;

function TAddIn.RegisterExtensionAs(var ExtensionName: String): Boolean;
begin
  Result := ClassFactory.RegisterExtensionAs(ExtensionName);
end;

function TAddIn.GetNProps: Integer;
begin
  Result := ClassFactory.GetNProps();
end;

function TAddIn.FindProp(const PropName: String): Integer;
begin
  Result := ClassFactory.FindProp(PropName);
end;

function TAddIn.GetPropName(const PropNum, PropAlias: Integer): String;
begin
  Result := ClassFactory.GetPropName(PropNum, PropAlias);
end;

function TAddIn.GetPropVal(const PropNum: Integer; var Value: Variant): Boolean;
begin
  Result := False;
  try
    Result := ClassFactory.GetPropVal(Self, PropNum, Value);
  except
    on E: Exception do
      AddError(ADDIN_E_FAIL, ClassName, E.Message, 0);
  end;
end;

function TAddIn.SetPropVal(const PropNum: Integer; const Value: Variant
  ): Boolean;
begin
  Result := False;
  try
    Result := ClassFactory.SetPropVal(Self, PropNum, Value);
  except
    on E: Exception do
      AddError(ADDIN_E_FAIL, ClassName, E.Message, 0);
  end;
end;

function TAddIn.IsPropReadable(const PropNum: Integer): Boolean;
begin
  Result := ClassFactory.IsPropReadable(PropNum);
end;

function TAddIn.IsPropWritable(const PropNum: Integer): Boolean;
begin
  Result := ClassFactory.IsPropWritable(PropNum);
end;

function TAddIn.GetNMethods: Integer;
begin
  Result := ClassFactory.GetNMethods;
end;

function TAddIn.FindMethod(const AMethodName: String): Integer;
begin
  Result := ClassFactory.FindMethod(AMethodName);
end;

function TAddIn.GetMethodName(const MethodNum, MethodAlias: Integer): String;
begin
  Result := ClassFactory.GetMethodName(MethodNum, MethodAlias);
end;

function TAddIn.GetNParams(const MethodNum: Integer): Integer;
begin
  Result := 0;
  try
    Result := ClassFactory.GetNParams(MethodNum);
  except
    on E: Exception do
      AddError(ADDIN_E_FAIL, ClassName, E.Message, 0);
  end;
end;

function TAddIn.GetParamDefValue(const MethodNum, ParamNum: Integer;
  var Value: Variant): Boolean;
begin
  Unused(MethodNum);
  Unused(ParamNum);
  Unused(Value);
  Result := False;
end;

function TAddIn.HasRetVal(const MethodNum: Integer): Boolean;
begin
  Result := False;
  try
    Result := ClassFactory.HasRetVal(MethodNum);
  except
    on E: Exception do
      AddError(ADDIN_E_FAIL, ClassName, E.Message, 0);
  end;
end;

function TAddIn.CallAsProc(const MethodNum: Integer;
  const Params: array of Variant): Boolean;
begin
  Result := False;
  try
    Result := ClassFactory.CallAsProc(Self, MethodNum, Params);
  except
    on E: Exception do
      AddError(ADDIN_E_FAIL, ClassName, E.Message, 0);
  end;
end;

function TAddIn.CallAsFunc(const MethodNum: Integer; var RetValue: Variant;
  const Params: array of Variant): Boolean;
begin
  Result := False;
  try
    Result := ClassFactory.CallAsFunc(Self, MethodNum, RetValue, Params);
  except
    on E: Exception do
      AddError(ADDIN_E_FAIL, ClassName, E.Message, 0);
  end;
end;

procedure TAddIn.SetLocale(const Locale: String);
begin
  Unused(Locale);
  // Do noting by default
end;

{$define ClassFactory := GetClassFactory(Self)}

class procedure TAddIn.RegisterAddInClass(const AddInExtensionName: String);
begin
  V8AddIn.RegisterAddInClass(Self);
  ClassFactory.AddInExtensionName := AddInExtensionName;
end;

class procedure TAddIn.AddProp(const Name, NameAlias: String;
  const ReadMethod: TGetMethod; const WriteMethod: TSetMethod);
begin
  ClassFactory.AddProp(Name, NameAlias, ReadMethod, WriteMethod);
end;

class procedure TAddIn.AddProc(const Name, NameAlias: String;
  const Method: TProcMethod; const NParams: Integer);
begin
  ClassFactory.AddProc(Name, NameAlias, Method, NParams);
end;

class procedure TAddIn.AddProc(const Name, NameAlias: String;
  const Method: TProc0Method);
begin
  ClassFactory.AddProc(Name, NameAlias, Method);
end;

class procedure TAddIn.AddProc(const Name, NameAlias: String;
  const Method: TProc1Method);
begin
  ClassFactory.AddProc(Name, NameAlias, Method);
end;

class procedure TAddIn.AddProc(const Name, NameAlias: String;
  const Method: TProc2Method);
begin
  ClassFactory.AddProc(Name, NameAlias, Method);
end;

class procedure TAddIn.AddProc(const Name, NameAlias: String;
  const Method: TProc3Method);
begin
  ClassFactory.AddProc(Name, NameAlias, Method);
end;

class procedure TAddIn.AddProc(const Name, NameAlias: String;
  const Method: TProc4Method);
begin
  ClassFactory.AddProc(Name, NameAlias, Method);
end;

class procedure TAddIn.AddProc(const Name, NameAlias: String;
  const Method: TProc5Method);
begin
  ClassFactory.AddProc(Name, NameAlias, Method);
end;

class procedure TAddIn.AddFunc(const Name, NameAlias: String;
  const Method: TFuncMethod; const NParams: Integer);
begin
  ClassFactory.AddFunc(Name, NameAlias, Method, NParams);
end;

class procedure TAddIn.AddFunc(const Name, NameAlias: String;
  const Method: TFunc0Method);
begin
  ClassFactory.AddFunc(Name, NameAlias, Method);
end;

class procedure TAddIn.AddFunc(const Name, NameAlias: String;
  const Method: TFunc1Method);
begin
  ClassFactory.AddFunc(Name, NameAlias, Method);
end;

class procedure TAddIn.AddFunc(const Name, NameAlias: String;
  const Method: TFunc2Method);
begin
  ClassFactory.AddFunc(Name, NameAlias, Method);
end;

class procedure TAddIn.AddFunc(const Name, NameAlias: String;
  const Method: TFunc3Method);
begin
  ClassFactory.AddFunc(Name, NameAlias, Method);
end;

class procedure TAddIn.AddFunc(const Name, NameAlias: String;
  const Method: TFunc4Method);
begin
  ClassFactory.AddFunc(Name, NameAlias, Method);
end;

class procedure TAddIn.AddFunc(const Name, NameAlias: String;
  const Method: TFunc5Method);
begin
  ClassFactory.AddFunc(Name, NameAlias, Method);
end;

{$undef ClassFactory}

{$POP}

{$I init.inc}

initialization

  InitUnit;

finalization

  ClearUnit;

end.

