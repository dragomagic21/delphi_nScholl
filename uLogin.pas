unit uLogin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, AdvGlassButton, StdCtrls, AdvEdit, AdvOfficePager;

type
  TfLogin = class(TForm)
    Pager: TAdvOfficePager;
    Page: TAdvOfficePage;
    eLogin: TAdvEdit;
    ePass: TAdvEdit;
    lPass: TLabel;
    lLogin: TLabel;
    bOk: TAdvGlassButton;
    bExit: TAdvGlassButton;
    procedure FormCreate(Sender: TObject);
    procedure bOkClick(Sender: TObject);
    procedure bExitClick(Sender: TObject);
    procedure ePassKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure eLoginKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    procedure CreateParams(var Params: TCreateParams); override;
  public
    { Public declarations }
  end;

var
  fLogin: TfLogin;

implementation

uses uData, uMain;

{$R *.dfm}

procedure TfLogin.CreateParams(var Params: TCreateParams);
begin
inherited CreateParams(Params);
with params do
   ExStyle := ExStyle or WS_EX_APPWINDOW;
end;

procedure TfLogin.FormCreate(Sender: TObject);
begin
 fData.SetLanguage(fLogin);
end;

procedure TfLogin.bOkClick(Sender: TObject);
var i: integer;
begin
 i := StrToInt(fData.cSelectS('ID','TB_USERS','where (LOGIN='+''''+Trim(ELogin.Text)+''''+')and(PASS='+''''+Trim(EPass.Text)+''''+')'));
 if ((i = -1)or(i = 0)) then fData.ShowError(18) else
 begin
  fData.UserID := i;
  fData.Admin := StrToInt(fData.cSelectS('RULES','TB_USERS','where ID='+IntToStr(i)));
  fMain.Mode := 1;
  fLogin.Close();
 end;
end;

procedure TfLogin.bExitClick(Sender: TObject);
begin
 Application.Terminate;
end;

procedure TfLogin.ePassKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
 if (Key = VK_RETURN) then bOkClick(self);
 if (Key = VK_ESCAPE) then bExitClick(self);
end;

procedure TfLogin.eLoginKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
 if (Key = VK_RETURN) then bOkClick(self);
 if (Key = VK_ESCAPE) then bExitClick(self);
end;

procedure TfLogin.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 Action := caFree;
end;

end.
