type

  { TPropData }

  TPropData = class
    Name: String;
    NameAlias: String;
    Read: TMethod;
    Write: TMethod;
    function DoRead(const Instance: TAddIn; var Value: Variant): Boolean;
    function DoWrite(const Instance: TAddIn; const Value: Variant): Boolean;
  end;

  { TMethodData }

  TMethodData = class
    Name: String;
    NameAlias: String;
    Method: TMethod;
    function GetNParams: Integer; virtual; abstract;
    function HasRetVal: Boolean; virtual; abstract;
    function CallAsProc(const Instance: TAddIn; const Params: array of Variant): Boolean; virtual; abstract;
    function CallAsFunc(const Instance: TAddIn; var RetValue: Variant; const Params: array of Variant): Boolean; virtual; abstract;
  end;

  { TCommonProcData }

  TCommonProcData = class (TMethodData)
    function HasRetVal: Boolean; override;
    function CallAsFunc(const Instance: TAddIn; var RetValue: Variant; const Params: array of Variant): Boolean; override;
  end;

  { TProcData }

  TProcData = class (TCommonProcData)
    NParams: Integer;
    function GetNParams: Integer; override;
    function CallAsProc(const Instance: TAddIn; const Params: array of Variant): Boolean; override;
  end;

  { TProc0Data }

  TProc0Data = class (TCommonProcData)
    function GetNParams: Integer; override;
    function CallAsProc(const Instance: TAddIn; const Params: array of Variant): Boolean; override;
  end;

  { TProc1Data }

  TProc1Data = class (TCommonProcData)
    function GetNParams: Integer; override;
    function CallAsProc(const Instance: TAddIn; const Params: array of Variant): Boolean; override;
  end;

  { TProc2Data }

  TProc2Data = class (TCommonProcData)
    function GetNParams: Integer; override;
    function CallAsProc(const Instance: TAddIn; const Params: array of Variant): Boolean; override;
  end;

  { TProc3Data }

  TProc3Data = class (TCommonProcData)
    function GetNParams: Integer; override;
    function CallAsProc(const Instance: TAddIn; const Params: array of Variant): Boolean; override;
  end;

  { TProc4Data }

  TProc4Data = class (TCommonProcData)
    function GetNParams: Integer; override;
    function CallAsProc(const Instance: TAddIn; const Params: array of Variant): Boolean; override;
  end;

  { TProc5Data }

  TProc5Data = class (TCommonProcData)
    function GetNParams: Integer; override;
    function CallAsProc(const Instance: TAddIn; const Params: array of Variant): Boolean; override;
  end;

  { TCommonFuncData }

  TCommonFuncData = class (TMethodData)
    function HasRetVal: Boolean; override;
    function CallAsProc(const Instance: TAddIn; const Params: array of Variant): Boolean; override;
  end;

  { TFuncData }

  TFuncData = class (TCommonFuncData)
    NParams: Integer;
    function GetNParams: Integer; override;
    function CallAsFunc(const Instance: TAddIn; var RetValue: Variant; const Params: array of Variant): Boolean; override;
  end;

  { TFunc0Data }

  TFunc0Data = class (TCommonFuncData)
    function GetNParams: Integer; override;
    function CallAsFunc(const Instance: TAddIn; var RetValue: Variant; const Params: array of Variant): Boolean; override;
  end;

  { TFunc1Data }

  TFunc1Data = class (TCommonFuncData)
    function GetNParams: Integer; override;
    function CallAsFunc(const Instance: TAddIn; var RetValue: Variant; const Params: array of Variant): Boolean; override;
  end;

  { TFunc2Data }

  TFunc2Data = class (TCommonFuncData)
    function GetNParams: Integer; override;
    function CallAsFunc(const Instance: TAddIn; var RetValue: Variant; const Params: array of Variant): Boolean; override;
  end;

  { TFunc3Data }

  TFunc3Data = class (TCommonFuncData)
    function GetNParams: Integer; override;
    function CallAsFunc(const Instance: TAddIn; var RetValue: Variant; const Params: array of Variant): Boolean; override;
  end;

  { TFunc4Data }

  TFunc4Data = class (TCommonFuncData)
    function GetNParams: Integer; override;
    function CallAsFunc(const Instance: TAddIn; var RetValue: Variant; const Params: array of Variant): Boolean; override;
  end;

  { TFunc5Data }

  TFunc5Data = class (TCommonFuncData)
    function GetNParams: Integer; override;
    function CallAsFunc(const Instance: TAddIn; var RetValue: Variant; const Params: array of Variant): Boolean; override;
  end;

  { TAddInFactory }

  TAddInFactory = class
    AddInClass: TAddInClass;
    AddInClassName: String;
    AddInClassVersion: Integer;
    PropList: TFPObjectList;
    MethodList: TFPObjectList;
    destructor Destroy; override;
    { Create destroy objects }
    function New: PComponentBase;
    procedure Delete(var AddInData: PComponentBase);
    { Easy use }
    function RegisterExtensionAs(var ExtensionName: String): Boolean;
    function GetNProps: Integer;
    function FindProp(const PropName: String): Integer;
    function GetPropName(const PropNum, PropAlias: Integer): String;
    function GetPropVal(const Instance: TAddIn; const PropNum: Integer; var Value: Variant): Boolean;
    function SetPropVal(const Instance: TAddIn; const PropNum: Integer; const Value: Variant): Boolean;
    function IsPropReadable(const PropNum: Integer): Boolean;
    function IsPropWritable(const PropNum: Integer): Boolean;
    function GetNMethods: Integer;
    function FindMethod(const AMethodName: String): Integer;
    function GetMethodName(const MethodNum, MethodAlias: Integer): String;
    function GetNParams(const MethodNum: Integer): Integer;
    function HasRetVal(const MethodNum: Integer): Boolean;
    function CallAsProc(const Instance: TAddIn; const MethodNum: Integer; const Params: array of Variant): Boolean;
    function CallAsFunc(const Instance: TAddIn; const MethodNum: Integer; var RetValue: Variant; const Params: array of Variant): Boolean;
    { Easy use initialization }
    procedure AddProp(const Name, NameAlias: String; const ReadMethod: TAddIn.TGetMethod; const WriteMethod: TAddIn.TSetMethod);
    procedure AddProc(const Name, NameAlias: String; const Method: TAddIn.TProcMethod; const NParams: Integer); overload;
    procedure AddProc(const Name, NameAlias: String; const Method: TAddIn.TProc0Method); overload;
    procedure AddProc(const Name, NameAlias: String; const Method: TAddIn.TProc1Method); overload;
    procedure AddProc(const Name, NameAlias: String; const Method: TAddIn.TProc2Method); overload;
    procedure AddProc(const Name, NameAlias: String; const Method: TAddIn.TProc3Method); overload;
    procedure AddProc(const Name, NameAlias: String; const Method: TAddIn.TProc4Method); overload;
    procedure AddProc(const Name, NameAlias: String; const Method: TAddIn.TProc5Method); overload;
    procedure AddFunc(const Name, NameAlias: String; const Method: TAddIn.TFuncMethod; const NParams: Integer); overload;
    procedure AddFunc(const Name, NameAlias: String; const Method: TAddIn.TFunc0Method); overload;
    procedure AddFunc(const Name, NameAlias: String; const Method: TAddIn.TFunc1Method); overload;
    procedure AddFunc(const Name, NameAlias: String; const Method: TAddIn.TFunc2Method); overload;
    procedure AddFunc(const Name, NameAlias: String; const Method: TAddIn.TFunc3Method); overload;
    procedure AddFunc(const Name, NameAlias: String; const Method: TAddIn.TFunc4Method); overload;
    procedure AddFunc(const Name, NameAlias: String; const Method: TAddIn.TFunc5Method); overload;
  end;

{ TPropData }

function TPropData.DoRead(const Instance: TAddIn; var Value: Variant): Boolean;
var
  ReadMethodRec: TMethod;
  ReadMethod: TAddIn.TGetMethod absolute ReadMethodRec;
begin
  ReadMethodRec := Read;
  ReadMethodRec.Data := Instance;
  Result := False;
  if not Assigned(ReadMethod) then
    Exit;
  Value := ReadMethod();
  Result := True;
end;

function TPropData.DoWrite(const Instance: TAddIn; const Value: Variant
  ): Boolean;
var
  WriteMethodRec: TMethod;
  WriteMethod: TAddIn.TSetMethod absolute WriteMethodRec;
begin
  WriteMethodRec := Write;
  WriteMethodRec.Data := Instance;
  Result := False;
  if not Assigned(WriteMethod) then
    Exit;
  WriteMethod(Value);
  Result := True;
end;

{ TCommonProcData }

function TCommonProcData.HasRetVal: Boolean;
begin
  Result := False;
end;

function TCommonProcData.CallAsFunc(const Instance: TAddIn;
  var RetValue: Variant; const Params: array of Variant): Boolean;
begin
  RetValue := Unassigned;
  Result := CallAsProc(Instance, Params);
end;

{ TProcData }

function TProcData.GetNParams: Integer;
begin
  Result := NParams;
end;

function TProcData.CallAsProc(const Instance: TAddIn;
  const Params: array of Variant): Boolean;
var
  MethodRec: TMethod;
  AMethod: TAddIn.TProcMethod absolute MethodRec;
begin
  MethodRec := Method;
  MethodRec.Data := Instance;
  Result := False;
  if not Assigned(AMethod) then
    Exit;
  AMethod(Params);
  Result := True;
end;

{ TProc0Data }

function TProc0Data.GetNParams: Integer;
begin
  Result := 0;
end;

function TProc0Data.CallAsProc(const Instance: TAddIn;
  const Params: array of Variant): Boolean;
var
  MethodRec: TMethod;
  AMethod: TAddIn.TProc0Method absolute MethodRec;
begin
  Unused(Params);
  MethodRec := Method;
  MethodRec.Data := Instance;
  Result := False;
  if not Assigned(AMethod) then
    Exit;
  AMethod();
  Result := True;
end;

{ TProc1Data }

function TProc1Data.GetNParams: Integer;
begin
  Result := 1;
end;

function TProc1Data.CallAsProc(const Instance: TAddIn;
  const Params: array of Variant): Boolean;
var
  MethodRec: TMethod;
  AMethod: TAddIn.TProc1Method absolute MethodRec;
begin
  MethodRec := Method;
  MethodRec.Data := Instance;
  Result := False;
  if not Assigned(AMethod) then
    Exit;
  AMethod(Params[0]);
  Result := True;
end;

{ TProc2Data }

function TProc2Data.GetNParams: Integer;
begin
  Result := 2;
end;

function TProc2Data.CallAsProc(const Instance: TAddIn;
  const Params: array of Variant): Boolean;
var
  MethodRec: TMethod;
  AMethod: TAddIn.TProc2Method absolute MethodRec;
begin
  MethodRec := Method;
  MethodRec.Data := Instance;
  Result := False;
  if not Assigned(AMethod) then
    Exit;
  AMethod(Params[0], Params[1]);
  Result := True;
end;

{ TProc3Data }

function TProc3Data.GetNParams: Integer;
begin
  Result := 3;
end;

function TProc3Data.CallAsProc(const Instance: TAddIn;
  const Params: array of Variant): Boolean;
var
  MethodRec: TMethod;
  AMethod: TAddIn.TProc3Method absolute MethodRec;
begin
  MethodRec := Method;
  MethodRec.Data := Instance;
  Result := False;
  if not Assigned(AMethod) then
    Exit;
  AMethod(Params[0], Params[1], Params[2]);
  Result := True;
end;

{ TProc4Data }

function TProc4Data.GetNParams: Integer;
begin
  Result := 4;
end;

function TProc4Data.CallAsProc(const Instance: TAddIn;
  const Params: array of Variant): Boolean;
var
  MethodRec: TMethod;
  AMethod: TAddIn.TProc4Method absolute MethodRec;
begin
  MethodRec := Method;
  MethodRec.Data := Instance;
  Result := False;
  if not Assigned(AMethod) then
    Exit;
  AMethod(Params[0], Params[1], Params[2], Params[3]);
  Result := True;
end;

{ TProc5Data }

function TProc5Data.GetNParams: Integer;
begin
  Result := 5;
end;

function TProc5Data.CallAsProc(const Instance: TAddIn;
  const Params: array of Variant): Boolean;
var
  MethodRec: TMethod;
  AMethod: TAddIn.TProc5Method absolute MethodRec;
begin
  MethodRec := Method;
  MethodRec.Data := Instance;
  Result := False;
  if not Assigned(AMethod) then
    Exit;
  AMethod(Params[0], Params[1], Params[2], Params[3], Params[4]);
  Result := True;
end;

{ TCommonFuncData }

function TCommonFuncData.HasRetVal: Boolean;
begin
  Result := True;
end;

function TCommonFuncData.CallAsProc(const Instance: TAddIn;
  const Params: array of Variant): Boolean;
var
  RetValue: Variant;
begin
  RetValue := Unassigned;
  Result := CallAsFunc(Instance, RetValue, Params);
end;

{ TFuncData }

function TFuncData.GetNParams: Integer;
begin
  Result := NParams;
end;

function TFuncData.CallAsFunc(const Instance: TAddIn; var RetValue: Variant;
  const Params: array of Variant): Boolean;
var
  MethodRec: TMethod;
  AMethod: TAddIn.TFuncMethod absolute MethodRec;
begin
  MethodRec := Method;
  MethodRec.Data := Instance;
  Result := False;
  if not Assigned(AMethod) then
    Exit;
  RetValue := AMethod(Params);
  Result := True;
end;

{ TFunc0Data }

function TFunc0Data.GetNParams: Integer;
begin
  Result := 0;
end;

function TFunc0Data.CallAsFunc(const Instance: TAddIn; var RetValue: Variant;
  const Params: array of Variant): Boolean;
var
  MethodRec: TMethod;
  AMethod: TAddIn.TFunc0Method absolute MethodRec;
begin
  Unused(Params);
  MethodRec := Method;
  MethodRec.Data := Instance;
  Result := False;
  if not Assigned(AMethod) then
    Exit;
  RetValue := AMethod();
  Result := True;
end;

{ TFunc1Data }

function TFunc1Data.GetNParams: Integer;
begin
  Result := 1;
end;

function TFunc1Data.CallAsFunc(const Instance: TAddIn; var RetValue: Variant;
  const Params: array of Variant): Boolean;
var
  MethodRec: TMethod;
  AMethod: TAddIn.TFunc1Method absolute MethodRec;
begin
  MethodRec := Method;
  MethodRec.Data := Instance;
  Result := False;
  if not Assigned(AMethod) then
    Exit;
  RetValue := AMethod(Params[0]);
  Result := True;
end;

{ TFunc2Data }

function TFunc2Data.GetNParams: Integer;
begin
  Result := 2;
end;

function TFunc2Data.CallAsFunc(const Instance: TAddIn; var RetValue: Variant;
  const Params: array of Variant): Boolean;
var
  MethodRec: TMethod;
  AMethod: TAddIn.TFunc2Method absolute MethodRec;
begin
  MethodRec := Method;
  MethodRec.Data := Instance;
  Result := False;
  if not Assigned(AMethod) then
    Exit;
  RetValue := AMethod(Params[0], Params[1]);
  Result := True;
end;

{ TFunc3Data }

function TFunc3Data.GetNParams: Integer;
begin
  Result := 3;
end;

function TFunc3Data.CallAsFunc(const Instance: TAddIn; var RetValue: Variant;
  const Params: array of Variant): Boolean;
var
  MethodRec: TMethod;
  AMethod: TAddIn.TFunc3Method absolute MethodRec;
begin
  MethodRec := Method;
  MethodRec.Data := Instance;
  Result := False;
  if not Assigned(AMethod) then
    Exit;
  RetValue := AMethod(Params[0], Params[1], Params[2]);
  Result := True;
end;

{ TFunc4Data }

function TFunc4Data.GetNParams: Integer;
begin
  Result := 4;
end;

function TFunc4Data.CallAsFunc(const Instance: TAddIn; var RetValue: Variant;
  const Params: array of Variant): Boolean;
var
  MethodRec: TMethod;
  AMethod: TAddIn.TFunc4Method absolute MethodRec;
begin
  MethodRec := Method;
  MethodRec.Data := Instance;
  Result := False;
  if not Assigned(AMethod) then
    Exit;
  RetValue := AMethod(Params[0], Params[1], Params[2], Params[3]);
  Result := True;
end;

{ TFunc5Data }

function TFunc5Data.GetNParams: Integer;
begin
  Result := 5;
end;

function TFunc5Data.CallAsFunc(const Instance: TAddIn; var RetValue: Variant;
  const Params: array of Variant): Boolean;
var
  MethodRec: TMethod;
  AMethod: TAddIn.TFunc5Method absolute MethodRec;
begin
  MethodRec := Method;
  MethodRec.Data := Instance;
  Result := False;
  if not Assigned(AMethod) then
    Exit;
  RetValue := AMethod(Params[0], Params[1], Params[2], Params[3], Params[4]);
  Result := True;
end;

{ TAddInFactory }

destructor TAddInFactory.Destroy;
begin
  if Assigned(AddInClass) then
    AddInClass.FFactory := nil;
  if PropList <> nil then
    FreeAndNil(PropList);
  if MethodList <> nil then
    FreeAndNil(MethodList);
  inherited Destroy;
end;

function TAddInFactory.New: PComponentBase;
begin
  Result := GetMem(SizeOf(TComponentBase));
  Result^ := ComponentBase;
  Result^.Factory := Self;
  Result^.Data := AddInClass.Create;
end;

procedure TAddInFactory.Delete(var AddInData: PComponentBase);
begin
  AddInData^.Data.Free;
  FreeMem(AddInData);
  AddInData := nil;
end;

function TAddInFactory.RegisterExtensionAs(var ExtensionName: String): Boolean;
begin
  ExtensionName := AddInClassName;
  Result := True;
end;

function TAddInFactory.GetNProps: Integer;
begin
  if PropList = nil then
    Result := 0
  else
    Result := PropList.Count;
end;

function TAddInFactory.FindProp(const PropName: String): Integer;
begin
  if PropList <> nil then
    begin
      for Result := 0 to PropList.Count - 1 do
        if AnsiCompareText(TPropData(PropList[Result]).Name, PropName) = 0 then
          Exit;
      for Result := 0 to PropList.Count - 1 do
        if AnsiCompareText(TPropData(PropList[Result]).NameAlias, PropName) = 0 then
          Exit;
    end;
  Result := -1;
end;

function TAddInFactory.GetPropName(const PropNum, PropAlias: Integer): String;
begin
  Result := '';
  if (PropNum < 0) or (PropList = nil) or (PropNum >= PropList.Count) then
    Exit;
  case PropAlias of
    0:
      Result := TPropData(PropList[PropNum]).Name;
    1:
      Result := TPropData(PropList[PropNum]).NameAlias;
  end;
end;

function TAddInFactory.GetPropVal(const Instance: TAddIn;
  const PropNum: Integer; var Value: Variant): Boolean;
begin
  Result := False;
  if (PropNum < 0) or (PropList = nil) or (PropNum >= PropList.Count) then
    Exit;
  Result := TPropData(PropList[PropNum]).DoRead(Instance, Value);
end;

function TAddInFactory.SetPropVal(const Instance: TAddIn;
  const PropNum: Integer; const Value: Variant): Boolean;
begin
  Result := False;
  if (PropNum < 0) or (PropList = nil) or (PropNum >= PropList.Count) then
    Exit;
  Result := TPropData(PropList[PropNum]).DoWrite(Instance, Value);
end;

function TAddInFactory.IsPropReadable(const PropNum: Integer): Boolean;
begin
  Result := False;
  if (PropNum < 0) or (PropList = nil) or (PropNum >= PropList.Count) then
    Exit;
  Result := Assigned(TAddIn.TGetMethod(TPropData(PropList[PropNum]).Read));
end;

function TAddInFactory.IsPropWritable(const PropNum: Integer): Boolean;
begin
  Result := False;
  if (PropNum < 0) or (PropList = nil) or (PropNum >= PropList.Count) then
    Exit;
  Result := Assigned(TAddIn.TSetMethod(TPropData(PropList[PropNum]).Write));
end;

function TAddInFactory.GetNMethods: Integer;
begin
  if MethodList = nil then
    Result := 0
  else
    Result := MethodList.Count;
end;

function TAddInFactory.FindMethod(const AMethodName: String): Integer;
begin
  if MethodList <> nil then
    begin
      for Result := 0 to MethodList.Count - 1 do
        if AnsiCompareText(TMethodData(MethodList[Result]).Name, AMethodName) = 0 then
          Exit;
      for Result := 0 to MethodList.Count - 1 do
        if AnsiCompareText(TMethodData(MethodList[Result]).NameAlias, AMethodName) = 0 then
          Exit;
    end;
  Result := -1;
end;

function TAddInFactory.GetMethodName(const MethodNum, MethodAlias: Integer
  ): String;
begin
  Result := '';
  if (MethodNum < 0) or (MethodList = nil) or (MethodNum >= MethodList.Count) then
    Exit;
  case MethodAlias of
    0:
      Result := TMethodData(MethodList[MethodNum]).Name;
    1:
      Result := TMethodData(MethodList[MethodNum]).NameAlias;
  end;
end;

function TAddInFactory.GetNParams(const MethodNum: Integer): Integer;
begin
  Result := 0;
  if (MethodNum < 0) or (MethodList = nil) or (MethodNum >= MethodList.Count) then
    Exit;
  Result := TMethodData(MethodList[MethodNum]).GetNParams;
end;

function TAddInFactory.HasRetVal(const MethodNum: Integer): Boolean;
begin
  Result := False;
  if (MethodNum < 0) or (MethodList = nil) or (MethodNum >= MethodList.Count) then
    Exit;
  Result := TMethodData(MethodList[MethodNum]).HasRetVal;
end;

function TAddInFactory.CallAsProc(const Instance: TAddIn;
  const MethodNum: Integer; const Params: array of Variant): Boolean;
begin
  Result := False;
  if (MethodNum < 0) or (MethodList = nil) or (MethodNum >= MethodList.Count) then
    Exit;
  Result := TMethodData(MethodList[MethodNum]).CallAsProc(Instance, Params);
end;

function TAddInFactory.CallAsFunc(const Instance: TAddIn;
  const MethodNum: Integer; var RetValue: Variant;
  const Params: array of Variant): Boolean;
begin
  Result := False;
  if (MethodNum < 0) or (MethodList = nil) or (MethodNum >= MethodList.Count) then
    Exit;
  Result := TMethodData(MethodList[MethodNum]).CallAsFunc(Instance, RetValue, Params);
end;

procedure TAddInFactory.AddProp(const Name, NameAlias: String;
  const ReadMethod: TAddIn.TGetMethod; const WriteMethod: TAddIn.TSetMethod);
var
  PropData: TPropData;
begin
  if PropList = nil then
    PropList := TFPObjectList.Create(True);
  PropData := TPropdata.Create;
  PropData.Name := Name;
  PropData.NameAlias := NameAlias;
  PropData.Read := TMethod(ReadMethod);
  PropData.Write := TMethod(WriteMethod);
  PropList.Add(PropData);
end;

procedure TAddInFactory.AddProc(const Name, NameAlias: String;
  const Method: TAddIn.TProcMethod; const NParams: Integer);
var
  MethodData: TProcData;
begin
  if MethodList = nil then
    MethodList := TFPObjectList.Create(True);
  MethodData := TProcData.Create;
  MethodData.Name := Name;
  MethodData.NameAlias := NameAlias;
  MethodData.Method := TMethod(Method);
  MethodData.NParams := NParams;
  MethodList.Add(MethodData);
end;

procedure TAddInFactory.AddProc(const Name, NameAlias: String;
  const Method: TAddIn.TProc0Method);
var
  MethodData: TProc0Data;
begin
  if MethodList = nil then
    MethodList := TFPObjectList.Create(True);
  MethodData := TProc0Data.Create;
  MethodData.Name := Name;
  MethodData.NameAlias := NameAlias;
  MethodData.Method := TMethod(Method);
  MethodList.Add(MethodData);
end;

procedure TAddInFactory.AddProc(const Name, NameAlias: String;
  const Method: TAddIn.TProc1Method);
var
  MethodData: TProc1Data;
begin
  if MethodList = nil then
    MethodList := TFPObjectList.Create(True);
  MethodData := TProc1Data.Create;
  MethodData.Name := Name;
  MethodData.NameAlias := NameAlias;
  MethodData.Method := TMethod(Method);
  MethodList.Add(MethodData);
end;

procedure TAddInFactory.AddProc(const Name, NameAlias: String;
  const Method: TAddIn.TProc2Method);
var
  MethodData: TProc2Data;
begin
  if MethodList = nil then
    MethodList := TFPObjectList.Create(True);
  MethodData := TProc2Data.Create;
  MethodData.Name := Name;
  MethodData.NameAlias := NameAlias;
  MethodData.Method := TMethod(Method);
  MethodList.Add(MethodData);
end;

procedure TAddInFactory.AddProc(const Name, NameAlias: String;
  const Method: TAddIn.TProc3Method);
var
  MethodData: TProc3Data;
begin
  if MethodList = nil then
    MethodList := TFPObjectList.Create(True);
  MethodData := TProc3Data.Create;
  MethodData.Name := Name;
  MethodData.NameAlias := NameAlias;
  MethodData.Method := TMethod(Method);
  MethodList.Add(MethodData);
end;

procedure TAddInFactory.AddProc(const Name, NameAlias: String;
  const Method: TAddIn.TProc4Method);
var
  MethodData: TProc4Data;
begin
  if MethodList = nil then
    MethodList := TFPObjectList.Create(True);
  MethodData := TProc4Data.Create;
  MethodData.Name := Name;
  MethodData.NameAlias := NameAlias;
  MethodData.Method := TMethod(Method);
  MethodList.Add(MethodData);
end;

procedure TAddInFactory.AddProc(const Name, NameAlias: String;
  const Method: TAddIn.TProc5Method);
var
  MethodData: TProc5Data;
begin
  if MethodList = nil then
    MethodList := TFPObjectList.Create(True);
  MethodData := TProc5Data.Create;
  MethodData.Name := Name;
  MethodData.NameAlias := NameAlias;
  MethodData.Method := TMethod(Method);
  MethodList.Add(MethodData);
end;

procedure TAddInFactory.AddFunc(const Name, NameAlias: String;
  const Method: TAddIn.TFuncMethod; const NParams: Integer);
var
  MethodData: TFuncData;
begin
  if MethodList = nil then
    MethodList := TFPObjectList.Create(True);
  MethodData := TFuncData.Create;
  MethodData.Name := Name;
  MethodData.NameAlias := NameAlias;
  MethodData.Method := TMethod(Method);
  MethodData.NParams := NParams;
  MethodList.Add(MethodData);
end;

procedure TAddInFactory.AddFunc(const Name, NameAlias: String;
  const Method: TAddIn.TFunc0Method);
var
  MethodData: TFunc0Data;
begin
  if MethodList = nil then
    MethodList := TFPObjectList.Create(True);
  MethodData := TFunc0Data.Create;
  MethodData.Name := Name;
  MethodData.NameAlias := NameAlias;
  MethodData.Method := TMethod(Method);
  MethodList.Add(MethodData);
end;

procedure TAddInFactory.AddFunc(const Name, NameAlias: String;
  const Method: TAddIn.TFunc1Method);
var
  MethodData: TFunc1Data;
begin
  if MethodList = nil then
    MethodList := TFPObjectList.Create(True);
  MethodData := TFunc1Data.Create;
  MethodData.Name := Name;
  MethodData.NameAlias := NameAlias;
  MethodData.Method := TMethod(Method);
  MethodList.Add(MethodData);
end;

procedure TAddInFactory.AddFunc(const Name, NameAlias: String;
  const Method: TAddIn.TFunc2Method);
var
  MethodData: TFunc2Data;
begin
  if MethodList = nil then
    MethodList := TFPObjectList.Create(True);
  MethodData := TFunc2Data.Create;
  MethodData.Name := Name;
  MethodData.NameAlias := NameAlias;
  MethodData.Method := TMethod(Method);
  MethodList.Add(MethodData);
end;

procedure TAddInFactory.AddFunc(const Name, NameAlias: String;
  const Method: TAddIn.TFunc3Method);
var
  MethodData: TFunc3Data;
begin
  if MethodList = nil then
    MethodList := TFPObjectList.Create(True);
  MethodData := TFunc3Data.Create;
  MethodData.Name := Name;
  MethodData.NameAlias := NameAlias;
  MethodData.Method := TMethod(Method);
  MethodList.Add(MethodData);
end;

procedure TAddInFactory.AddFunc(const Name, NameAlias: String;
  const Method: TAddIn.TFunc4Method);
var
  MethodData: TFunc4Data;
begin
  if MethodList = nil then
    MethodList := TFPObjectList.Create(True);
  MethodData := TFunc4Data.Create;
  MethodData.Name := Name;
  MethodData.NameAlias := NameAlias;
  MethodData.Method := TMethod(Method);
  MethodList.Add(MethodData);
end;

procedure TAddInFactory.AddFunc(const Name, NameAlias: String;
  const Method: TAddIn.TFunc5Method);
var
  MethodData: TFunc5Data;
begin
  if MethodList = nil then
    MethodList := TFPObjectList.Create(True);
  MethodData := TFunc5Data.Create;
  MethodData.Name := Name;
  MethodData.NameAlias := NameAlias;
  MethodData.Method := TMethod(Method);
  MethodList.Add(MethodData);
end;


