unit uClassRz;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, AdvEdit, Grids, BaseGrid, AdvGrid, StdCtrls, AdvFontCombo,
  AdvOfficePager, AdvGlassButton;

type
  TfClassRz = class(TForm)
    Pager: TAdvOfficePager;
    Page: TAdvOfficePage;
    eF: TAdvOfficeComboBox;
    lcase1: TLabel;
    lcase2: TLabel;
    sgClass: TAdvStringGrid;
    eNNum: TAdvEdit;
    eNLetter: TAdvEdit;
    sgNew: TAdvStringGrid;
    bIns: TAdvGlassButton;
    bDel: TAdvGlassButton;
    bOk: TAdvGlassButton;
    lReport: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure eFChange(Sender: TObject);
    procedure eNLetterChange(Sender: TObject);
    procedure bInsClick(Sender: TObject);
    procedure bDelClick(Sender: TObject);
    procedure bOkClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    Name: byte;
    p: array of integer;
    { Private declarations }
    procedure CreateParams(var Params: TCreateParams); override;
    // Добавление ИДшника
    function Ins(ID: integer): byte;
    // Удаление ИДшника
    function Del(num: integer): byte;
    // Заполнение списка учеников нового класса
    function FillNew(): byte;
    // Формируем строку из ИДшников через запятую
    function Exp(): string;
  public
    School: integer;
    { Public declarations }
  end;

var
  fClassRz: TfClassRz;

implementation

uses uData, uMain;

{$R *.dfm}

procedure TfClassRz.CreateParams(var Params: TCreateParams);
begin
inherited CreateParams(Params);
with params do
   ExStyle := ExStyle or WS_EX_APPWINDOW;
end;

procedure TfClassRz.FormCreate(Sender: TObject);
begin
 // Устанавливаем язык формы
 fData.SetLanguage(Self);
 // Устанавливаем заголовки
 sgClass.Cells[0,0] := '<P ALIGN="CENTER">'+fData.standarts[1]+'</P>';
 sgNew.Cells[0,0] := '<P ALIGN="CENTER">'+fData.standarts[1]+'</P>';
 // Обнуляем массив ИДшников
 SetLength(p, 0);
end;

procedure TfClassRz.FormShow(Sender: TObject);
begin
 // Заполняем список классов
 fData.FillClassCb(eF, 'Where ((SCHOOL='+IntToStr(School)+')and(NUM<'+IntToStr(fData.MaxClass)+'))');
 eFChange(Self);
end;

procedure TfClassRz.eFChange(Sender: TObject);
begin
 // Устанавливаем номер нового класса
 eNNum.Text := fData.cSelectS('NUM','TB_CLASS','where ID='+IntToStr(integer(eF.Items.Objects[eF.ItemIndex])));
 eNLetterChange(Self);
 // Заполняем учеников выбранного класса
 case fData.FillFIOSg('TB_PEOP', 'CLASS', integer(eF.Items.Objects[eF.ItemIndex]), sgClass) of
  0: bIns.Enabled := TRUE;
  1: bIns.Enabled := FALSE;
 end;
 // Обнуляем список ИДшников
 SetLength(p, 0);
 // Заполняем список нового класса
 if (FillNew() = 1) then bDel.Enabled := FALSE else bDel.Enabled := TRUE;
end;

procedure TfClassRz.eNLetterChange(Sender: TObject);
begin
 // Не пусто ли?
 if (Trim(eNLetter.Text) <> '') then
  // Проверяем есть ли такой класс уже
  if (fData.cCount('ID','TB_CLASS','where ((SCHOOL='+IntToStr(School)+')and(NUM='+eNNum.Text+')and(NAME='''+eNLetter.Text+'''))') = 0) then
  begin
   lReport.Caption := '';
   Name := 0;
  end else
  begin
   lReport.Caption := fData.standarts[23];
   Name := 1;
  end; 
end;

procedure TfClassRz.bInsClick(Sender: TObject);
begin
 // Добавляем ИДшник
 Ins(integer(sgClass.Objects[0, sgClass.Row]));
 // Выводим список выбранного класса с фильтром
 case fData.FillFIOSg('TB_PEOP', 'CLASS', integer(eF.Items.Objects[eF.ItemIndex]), sgClass, 'and(ID not in ('+Exp()+'))') of
  0: bIns.Enabled := TRUE;
  1: bIns.Enabled := FALSE;
 end;
 // Заполняем список учеников нового класса
 if (FillNew() = 1) then bDel.Enabled := FALSE else bDel.Enabled := TRUE;
 if ((sgClass.RowCount > 1)and(sgNew.RowCount > 1)and(Name = 0)) then bOk.Enabled := TRUE else bOk.Enabled := FALSE;
end;

function TfClassRz.FillNew: byte;
var i: integer;
begin
 sgNew.RowCount := 1;
 for i := 0 to Length(p)-1 do
 begin
  sgNew.RowCount := sgNew.RowCount+1;
  sgNew.Objects[0, sgNew.RowCount-1] := Pointer(p[i]);
  sgNew.Cells[0, sgNew.RowCount-1]:= fData.cSelectS('FNAME', 'TB_PEOP', 'where ID='+IntToStr(p[i])) + ' ' + fData.cSelectS('NAME', 'TB_PEOP', 'where ID='+IntToStr(p[i]))[1] + '. ' + fData.cSelectS('SNAME', 'TB_PEOP', 'where ID='+IntToStr(p[i]))[1] + '.';
 end;
 if (sgNew.RowCount > 1) then
 begin
  Result := 0;
  sgNew.FixedRows := 1;
 end else
 begin
  Result := 1;
  sgNew.FixedRows := 0;
 end;

end;

function TfClassRz.Ins(ID: integer): byte;
begin
 try
  Result := 0;
  SetLength(p, Length(p)+1);
  p[Length(p)-1] := ID;
 except
  Result := 1;
 end;
end;

function TfClassRz.Del(num: integer): byte;
var i: integer;
begin
 try
  Result := 0;
  for i := num to Length(p)-2 do
   p[i] := p[i+1];
  SetLength(p, Length(p)-1);
 except
  Result := 1;
 end;
end;

procedure TfClassRz.bDelClick(Sender: TObject);
begin
 Del(sgNew.Row-1);
 case fData.FillFIOSg('TB_PEOP', 'CLASS', integer(eF.Items.Objects[eF.ItemIndex]), sgClass, 'and(ID not in ('+Exp()+'))') of
  0: bIns.Enabled := TRUE;
  1: bIns.Enabled := FALSE;
 end;
 if (FillNew() = 1) then bDel.Enabled := FALSE else bDel.Enabled := TRUE;
 if ((sgClass.RowCount > 1)and(sgNew.RowCount > 1)and(Name = 0)) then bOk.Enabled := TRUE else bOk.Enabled := FALSE;
end;

function TfClassRz.Exp: string;
var tmp: string;
    i: integer;
begin
 tmp := '';
 for i := 0 to Length(p)-1 do
  tmp := tmp + IntToStr(p[i]) + ',';
 delete(tmp,Length(tmp),1);
 Result := tmp;
end;

procedure TfClassRz.bOkClick(Sender: TObject);
var i, nc: integer;
    r,v: TstringList;
begin
 if (Length(p) > 0) then
 begin
 // Добавляем новый класс
 r := TStringList.Create(); r.Clear();
 v := TStringList.Create(); v.Clear();
 r.Add('SCHOOL'); r.Add('NUM'); r.Add('NAME');
 v.Add(IntToStr(School)); v.Add(eNNum.Text); v.Add(''''+eNLetter.Text+'''');
 fData.cInserts('TB_CLASS',r,v);
 r.Free(); v.Free();
 nc := StrToInt(fData.cSelectS('ID','TB_CLASS','where ((SCHOOL='+IntToStr(School)+')and(NUM='+eNNum.Text+')and(NAME='''+eNLetter.Text+'''))'));
 // Бежим по отобранному списку и делаем АПДЭЙТ по полю КЛАСС
 for i := 0 to Length(p)-1 do
  fData.cUpdate('TB_PEOP','CLASS',IntToStr(nc),'where ID='+IntToStr(p[i]));
 fMain.sgSchoolsClick(Self);
 Self.Close();
 end else fData.ShowError(10); 
end;

procedure TfClassRz.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 Action := caFree;
end;

end.
