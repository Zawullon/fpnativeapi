unit V8AddIn;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, CTypes, Variants, LCLProc;

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

  TAddIn = class
  private
    FConnection: Pointer;
    FMemoryManager: Pointer;
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
    function Confirm(const QueryText: String; var RetValue: Variant): Boolean;
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
  end;

  TAddInClass = class of TAddIn;

procedure Unused(const AValue);

procedure RegisterAddInClass(const AddInClassName: String; const AddInClass: TAddInClass);

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

procedure RegisterAddInClass(const AddInClassName: String;
  const AddInClass: TAddInClass);
var
  Index: Integer;
  Factory: TAddInFactory;
begin
  Index := FactoryList.IndexOf(AddInClassName);
  if Index >= 0 then
    FactoryList.Delete(Index);
  Factory := TAddInFactory.Create;
  Factory.AddInClass := AddInClass;
  FactoryList.AddObject(AddInClassName, Factory);
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
  Result := 0001;
end;

procedure TAddIn.Done;
begin
  // Do noting by default
end;

function TAddIn.RegisterExtensionAs(var ExtensionName: String): Boolean;
begin
  Unused(ExtensionName);
  Result := False;
end;

function TAddIn.GetNProps: Integer;
begin
  Result := 0;
end;

function TAddIn.FindProp(const PropName: String): Integer;
begin
  Unused(PropName);
  Result := -1;
end;

function TAddIn.GetPropName(const PropNum, PropAlias: Integer): String;
begin
  Unused(PropNum);
  Unused(PropAlias);
  Result := '';
end;

function TAddIn.GetPropVal(const PropNum: Integer; var Value: Variant): Boolean;
begin
  Unused(PropNum);
  Unused(Value);
  Result := False;
end;

function TAddIn.SetPropVal(const PropNum: Integer; const Value: Variant
  ): Boolean;
begin
  Unused(PropNum);
  Unused(Value);
  Result := False;
end;

function TAddIn.IsPropReadable(const PropNum: Integer): Boolean;
begin
  Unused(PropNum);
  Result := False;
end;

function TAddIn.IsPropWritable(const PropNum: Integer): Boolean;
begin
  Unused(PropNum);
  Result := False;
end;

function TAddIn.GetNMethods: Integer;
begin
  Result := 0;
end;

function TAddIn.FindMethod(const AMethodName: String): Integer;
begin
  Unused(AMethodName);
  Result := -1;
end;

function TAddIn.GetMethodName(const MethodNum, MethodAlias: Integer): String;
begin
  Unused(MethodNum);
  Unused(MethodAlias);
  Result := '';
end;

function TAddIn.GetNParams(const MethodNum: Integer): Integer;
begin
  Unused(MethodNum);
  Result := 0;
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
  Unused(MethodNum);
  Result := False;
end;

function TAddIn.CallAsProc(const MethodNum: Integer;
  const Params: array of Variant): Boolean;
begin
  Unused(MethodNum);
  Unused(Params);
  Result := False;
end;

function TAddIn.CallAsFunc(const MethodNum: Integer; var RetValue: Variant;
  const Params: array of Variant): Boolean;
begin
  Unused(MethodNum);
  Unused(RetValue);
  Unused(Params);
  Result := False;
end;

procedure TAddIn.SetLocale(const Locale: String);
begin
  Unused(Locale);
  // Do noting by default
end;

{$I init.inc}

initialization

  InitUnit;

finalization

  ClearUnit;

end.

