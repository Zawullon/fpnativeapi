unit GUIExample;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, V8AddIn, V8AddInGUI, ExampleForm;

type

  { TFormController }

  TFormController = class (TAddIn)
  private
    FText: String;
    Form: TFormExample;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure DoInit;
    procedure DoDone;
    procedure DoSetText;
  public
    function Init: Boolean; override;
    procedure Done; override;
    procedure Show;
    procedure Hide;
    procedure SetText(Text: Variant);
  end;

implementation

{ TFormController }

procedure TFormController.Button1Click(Sender: TObject);
begin
  ExternalEvent('GUI', 'Event', 'Button 1 click');
end;

procedure TFormController.Button2Click(Sender: TObject);
begin
  ExternalEvent('GUI', 'Event', 'Button 2 click');
end;

procedure TFormController.DoInit;
begin
  Form := TFormExample.Create(nil);
  Form.Button1.OnClick := @Button1Click;
  Form.Button2.OnClick := @Button2Click;
end;

procedure TFormController.DoDone;
begin
  Form.Free;
end;

procedure TFormController.DoSetText;
begin
  Form.Label1.Caption := FText;
end;

function TFormController.Init: Boolean;
begin
  GUIThread.Synchronize(@DoInit);
  Result := True;
end;

procedure TFormController.Done;
begin
  GUIThread.Synchronize(@DoDone);
  inherited Done;
end;

procedure TFormController.Show;
begin
  GUIThread.Synchronize(@Form.Show);
end;

procedure TFormController.Hide;
begin
  GUIThread.Synchronize(@Form.Hide);
end;

procedure TFormController.SetText(Text: Variant);
begin
  FText := Text;
  GUIThread.Synchronize(@DoSetText);
end;

end.

