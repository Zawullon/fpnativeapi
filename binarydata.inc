{****************************************************************************
                        1C binary data
****************************************************************************}

type
  P1CBinaryData = ^T1CBinaryData;
  T1CBinaryData = record
    Size: SizeInt;
    Data: Byte;
  end;

{ T1CBinaryDataFactory }
type
  T1CBinaryDataFactory = class (TCustomVariantType)
  public
    procedure Clear(var V: TVarData); override;
    procedure Copy(var Dest: TVarData; const Source: TVarData; const Indirect: Boolean); override;
  end;

procedure T1CBinaryDataFactory.Clear(var V: TVarData);
begin
  FreeMem(V.VPointer);
  V.VType := varEmpty;
end;

procedure T1CBinaryDataFactory.Copy(var Dest: TVarData; const Source: TVarData;
  const Indirect: Boolean);
begin
  if Indirect then
    Dest.vpointer := Source.vpointer
  else
    begin
      Dest.vpointer := GetMem(P1CBinaryData(Source.vpointer)^.Size + SizeOf(SizeInt));
      Move(Source.vpointer^, Dest.vpointer^, P1CBinaryData(Source.vpointer)^.Size + SizeOf(SizeInt));
    end;
  Dest.VType := VarType;
end;

var
  BinaryDataFactory: T1CBinaryDataFactory;

function StreamTo1CBinaryData(const Stream: TStream): Variant;
var
  Data: P1CBinaryData;
begin
  Data := Getmem(Stream.Size + SizeOf(SizeInt));
  Data^.Size := Stream.Size;
  Stream.Position := 0;
  Stream.Read(Data^.Data, Data^.Size);
  TVarData(Result).vtype := BinaryDataFactory.VarType;
  TVarData(Result).vpointer := Data;
end;

function Var1CBinaryData: TVarType;
begin
  Result := BinaryDataFactory.VarType;
end;

function Is1CBinaryData(const Value: Variant): Boolean;
begin
  Result := (VarType(Value) = BinaryDataFactory.VarType);
end;

function Get1CBinaryDataSize(const Value: Variant): SizeInt;
begin
  if VarType(Value) = BinaryDataFactory.VarType then
    Result := P1CBinaryData(TVarData(Value).vpointer)^.Size
  else
    Result := 0;
end;

function Get1CBinaryData(const Value: Variant): Pointer;
begin
  if VarType(Value) = BinaryDataFactory.VarType then
    Result := @(P1CBinaryData(TVarData(Value).vpointer)^.Data)
  else
    Result := nil;
end;

