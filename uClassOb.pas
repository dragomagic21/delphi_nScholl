unit uClassOb;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, AdvGlassButton, StdCtrls, AdvFontCombo, AdvOfficePager;

type
  TfClassOb = class(TForm)
    Pager: TAdvOfficePager;
    Page: TAdvOfficePage;
    eF: TAdvOfficeComboBox;
    eS: TAdvOfficeComboBox;
    lcase1: TLabel;
    lcase2: TLabel;
    bOk: TAdvGlassButton;
    procedure bOkClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure eFChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    b: byte;
    { Private declarations }
    procedure CreateParams(var Params: TCreateParams); override;
  public
    School: integer;
    { Public declarations }
  end;

var
  fClassOb: TfClassOb;

implementation

uses uData, uMain;

{$R *.dfm}

procedure TfClassOb.bOkClick(Sender: TObject);
begin
 if (fData.ClassO(integer(eF.Items.Objects[eF.ItemIndex]), integer(eS.Items.Objects[eS.ItemIndex])) = 0) then
 begin
  fMain.sgSchoolsClick(Self);
  Self.Close();
 end;
end;

procedure TfClassOb.CreateParams(var Params: TCreateParams);
begin
inherited CreateParams(Params);
with params do
   ExStyle := ExStyle or WS_EX_APPWINDOW;
end;

procedure TfClassOb.FormShow(Sender: TObject);
begin
 b := fData.FillClassCb(eF, 'Where ((SCHOOL='+IntToStr(School)+')and(NUM<'+IntToStr(fData.MaxClass)+'))');
 eFChange(Self);
end;

procedure TfClassOb.eFChange(Sender: TObject);
begin
 if (fData.FillClassCb(eS, 'Where ((SCHOOL='+IntToStr(School)+')and(NUM<'+IntToStr(fData.MaxClass)+')and(ID<>'+IntToStr(integer(eF.Items.Objects[eF.ItemIndex]))+'))') = 1) then b := 1;
 case b of
  0: bOk.Enabled := TRUE;
  1: bOk.Enabled := FALSE;
 end;
end;

procedure TfClassOb.FormCreate(Sender: TObject);
begin
 fData.SetLanguage(Self);
end;

procedure TfClassOb.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 Action := caFree;
end;

end.
