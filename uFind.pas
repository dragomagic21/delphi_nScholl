unit uFind;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, BaseGrid, AdvGrid, AdvGlassButton, ExtCtrls, StdCtrls,
  AdvEdit, AdvOfficePager;

type
  TfFind = class(TForm)
    Pager: TAdvOfficePager;
    Page: TAdvOfficePage;
    EName: TAdvEdit;
    Bevel1: TBevel;
    bFind: TAdvGlassButton;
    sgFind: TAdvStringGrid;
    procedure ENameChange(Sender: TObject);
    procedure bFindClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure sgFindDblClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    procedure CreateParams(var Params: TCreateParams); override;
  public
    { Public declarations }
  end;

var
  fFind: TfFind;

implementation

uses uData, uMain;

{$R *.dfm}

procedure TfFind.ENameChange(Sender: TObject);
begin
 fData.Find(sgFind, Trim(EName.Text));
end;

procedure TfFind.bFindClick(Sender: TObject);
begin
 fData.Find(sgFind, Trim(EName.Text));
end;

procedure TfFind.CreateParams(var Params: TCreateParams);
begin
inherited CreateParams(Params);
with params do
   ExStyle := ExStyle or WS_EX_APPWINDOW;
end;

procedure TfFind.FormShow(Sender: TObject);
begin
 bFindClick(self);
end;

procedure TfFind.sgFindDblClick(Sender: TObject);
var i,z: integer;
begin
 if (sgFind.RowCount > 0) then
 begin
  z := 0;
  // ищем школу
  for i := 0 to fMain.tSchools.AdvOfficeTabs.Count-1 do
   if (fMain.tSchools.AdvOfficeTabs[i].Tag = integer(sgFind.Objects[1, sgFind.Row])) then
   begin
    z := i;
    break;
   end;
  fMain.tSchools.ActiveTabIndex := z;
  fMain.tSchoolsChange(Self);
  fMain.Pager.ActivePageIndex := 1;
  fMain.PagerChange(Self);
  fMain.pClassInfo.ActivePageIndex := 1;
  fMain.pClassInfoChange(Self);

  // ищем класс
  z := 1;
  for i := 1 to fMain.sgClassNames.RowCount-1 do
   if (integer(fMain.sgClassNames.Objects[0, i]) = integer(sgFind.Objects[2, sgFind.Row])) then
   begin
    z := i;
    break;
   end;
  fMain.sgClassNames.Row := z;
  fMain.sgClassNamesClick(Self);

  // ищем ученика
  z := 1;
  for i := 1 to fMain.sgClassInfo.RowCount-1 do
   if (integer(fMain.sgClassInfo.Objects[0, i]) = integer(sgFind.Objects[0, sgFind.Row])) then
   begin
    z := i;
    break;
   end;
  fMain.sgClassInfo.Row := z;
  fMain.sgClassInfoClick(Self);

 end;
end;

procedure TfFind.FormCreate(Sender: TObject);
begin
 fData.SetLanguage(Self);
end;

procedure TfFind.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 Action := caFree;
end;

end.
