unit uPlanImport;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, AdvOfficePager, Grids, BaseGrid, AdvGrid, AdvGlassButton,
  ExtCtrls, AdvPanel, AdvOfficeTabSet;

type
  TfPlanImport = class(TForm)
    Pager: TAdvOfficePager;
    Page: TAdvOfficePage;
    pClasses: TAdvOfficeTabSet;
    pNavigation: TAdvPanel;
    bCancel: TAdvGlassButton;
    bOk: TAdvGlassButton;
    sg: TAdvStringGrid;
    procedure FormCreate(Sender: TObject);
    procedure pClassesChange(Sender: TObject);
    procedure bCancelClick(Sender: TObject);
    procedure bOkClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    procedure CreateParams(var Params: TCreateParams); override;
  public
    School: integer;
    ID: integer;
    { Public declarations }
  end;

var
  fPlanImport: TfPlanImport;

implementation

uses uData, uMain;

{$R *.dfm}

procedure TfPlanImport.FormCreate(Sender: TObject);
begin
 fData.SetLanguage(Self);
end;

procedure TfPlanImport.pClassesChange(Sender: TObject);
begin
 case fData.FillPlanSg(pClasses.AdvOfficeTabs[pClasses.ActiveTabIndex].Tag,sg) of
  0: bOk.Enabled := TRUE;
  1: bOk.Enabled := FALSE;
 end;
end;

procedure TfPlanImport.bCancelClick(Sender: TObject);
begin
 Self.Close();
end;

procedure TfPlanImport.bOkClick(Sender: TObject);
var i: integer;
    r,v: TStringList;
    res: byte;
begin
 try
  // вначале очищаем список, который есть сейчас
  fData.cDelete('TB_PLANCLASS', 'where UPLAN='+IntToStr(ID));
  
  // добавляем новые данные
  res := 0;
  r := TStringList.Create(); r.Clear();
  v := TStringList.Create(); v.Clear();
  r.Add('UPLAN'); r.Add('TEACHER'); r.Add('PREDMET');
  try
   for i := 1 to sg.RowCount-1 do
   begin
    v.Add(IntToStr(ID)); v.Add(IntToStr(integer(sg.Objects[1, i]))); v.Add(IntToStr(integer(sg.Objects[2, i])));
    fData.cInserts('TB_PLANCLASS',r,v);
    v.Clear();
   end;
   res := 0;
  except
   res := 1;
   fData.ShowError(4);
  end;
 finally
  r.Free;
  v.Free;
  if (res = 0) then
  begin
   fMain.tPlansClassNameChange(Self);
   Self.Close();
  end;
 end;
end;

procedure TfPlanImport.CreateParams(var Params: TCreateParams);
begin
inherited CreateParams(Params);
with params do
   ExStyle := ExStyle or WS_EX_APPWINDOW;
end;

procedure TfPlanImport.FormShow(Sender: TObject);
begin
 case fData.FillTabsText(pClasses, 'NAME', 'TB_UPLAN', 'where ((SCHOOL='+IntToStr(School)+')and(ID <> '+IntToStr(ID)+'))', 'order by NAME') of
  0: pClassesChange(self);
  1: pClassesChange(self);
 end;
end;

procedure TfPlanImport.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
 Action := caFree;
end;

end.
