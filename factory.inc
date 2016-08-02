{ TAddInFactory }

type
  TAddInFactory = class public
    AddInClass: TAddInClass;
    function New: PComponentBase;
    procedure Delete(var AddInData: PComponentBase);
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


