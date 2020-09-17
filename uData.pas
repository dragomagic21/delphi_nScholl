unit uData;

interface

uses
  Windows, Consts, SysUtils, Classes, StdCtrls, Forms, IniFiles, Dialogs,
  AdvOfficePager, AdvOfficeTabSet, AdvPanel, AdvGlowButton, AdvGrid,
  AdvReflectionLabel, AdvGlassButton, AdvCombo, AdvFontCombo, AdvOfficeButtons,
  IBSQL, IBDatabase, DB, Menus, OleServer, Word2000, Excel2000, Controls, Math, ExtCtrls,
  jpeg, IBSQLMonitor, Variants, ComObj, WordXP, ExcelXP;

const
 Labels: array [0..25] of char = ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J',
                                  'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T',
                                  'U', 'V', 'W', 'X', 'Y', 'Z' );

type
  TfData = class(TDataModule)
    Database: TIBDatabase;
    Transaction: TIBTransaction;
    SQL: TIBSQL;
    ExcelA: TExcelApplication;
    SQLMonitor: TIBSQLMonitor;
    WordA: TWordApplication;
    procedure DataModuleCreate(Sender: TObject);
    procedure SQLMonitorSQL(EventText: String; EventTime: TDateTime);
  private
   LengthLog: int64;
   { Private declarations }
   function SetConfig(): byte;
   function GetLabel(Num: integer): string;
   function GetNormalName(s: string): string;
   function GetFileSize(Path: string): int64;
  public
   language: string;
   MaxClass: integer;
   errors: array of string;
   standarts: array of string;
   JumpClass: array of integer;
   UserID, Admin: integer;
   Term: byte;
   { Public declarations }
   function cCount(What, Table: string; Where: string = ''): integer;
   function cMax(What, Table: string; Where: string = ''): integer;
   function ifCount(What, Table: string; Where: string = ''): integer;
   function cSelectS(What, Table: string; Where: string = ''): string;
   function cInsert(Table, Rec, Val: string): byte;
   function cInserts(Table: string; Recs, Vals: TStringList): byte;
   function cUpdate(Table, Rec, Val: string; Where: string = ''): byte;
   function cUpdates(Table: string; Recs, Vals: TStringList; Where: string = ''): byte;
   function cDelete(Table: string; Where: string = ''): byte;
   function cFillCb(cb: TAdvOfficeComboBox; What, Table: string; Where: string = ''): byte;
   function cFillSg(sg: TAdvStringGrid; Er: integer; What, Table: string; Where: string = ''): byte;
   // выборка ФИО StringGrid
   function FillFIOSg(Table, Filter: string; ID: integer; sg: TAdvStringGrid; andWhere: string = ''; cwidth: integer = 110): byte;
   // выборка ФИО ComboBox
   function FillFIOCb(Table, Filter: string; ID: integer; cb: TAdvOfficeComboBox): byte;
   // установка языка для формы
   function SetLanguage(fm: TForm): byte;
   // сообщение об ошибке
   procedure ShowError(Error: integer; Handle: HWND = 0);
   // выборка школ
   function FillSchollsTab(Tab: TAdvOfficeTabSet): byte;
   function FillSchollsSG(SG: TAdvStringGrid): byte;
   // выборка классов для школы
   function FillClassTab(School: integer; Tab: TAdvOfficeTabSet; Where: string = ''): byte;
   // выборка классов для школы
   function FillClassSG(School: integer; sg: TAdvStringGrid; Where: string = ''): byte;
   // Заполнение классов в КомбоБокс
   function FillClassCb(cb: TAdvOfficeComboBox; Where: string = ''): byte;
   // Заполнение выпускных классов в Грид
   function FillClassEndSG(sg: TAdvStringGrid; Where: string = ''; order: string = ''): byte;
   // выборка учебного плана
   function FillPlanSg(ID: integer; sg: TAdvStringGrid; Order: string = 'order by TB_TEACHER.FNAME,TB_TEACHER.NAME,TB_TEACHER.SNAME'): byte;
   // Фильтр учебного плана
   function FillPlanSGf(tp: byte; School: Integer;  ID: Integer; sg: TAdvStringGrid): byte;
   // Изменение учебного плана для класса
   function ChangePlan(Clas, UPlan: integer): byte;
   // Заполнение списка предметов для учебного плана (за исключением тех, которые уже есть в плане)
   function FillCbPredmetPlan(cb: TAdvOfficeComboBox; ID: integer): byte;
   // Экспорт в Эксель (уч. планы)
   function ExpPlansExcel(tp: byte; ID: integer; sg: TAdvStringGrid): byte;
   // Экспорт в Эксель Оценки все на все
   function ExpMarksExcelAll(sg: TAdvStringGrid): byte;
   // Экспорт в Эксель Оценки (Ученики\Предметы)
   function ExpMarksExcelP(tp: byte; sg: TAdvStringGrid): byte;
   // Экспорт в Эксель Оценки (Диаграмма)
   function ExpMarksDiag(tp: byte; sg: TAdvStringGrid): byte;
   // Экспорт в Эксель (мед. книга)
   function ExpMedExcel(sg: TAdvStringGrid): byte;
   // Заполнение учеников для оценок
   function FillFIOMarksSG(ID: integer; sg: TAdvStringGrid; sgw: integer = 110): byte;
   // Заполнение предметов для оценок
   function FillPrMarksSG(ID: integer; sg: TAdvStringGrid): byte;
   // Заполнение оценок все на все
   function FillMarksAll(ID: integer; sg: TAdvStringGrid): byte;
   // Заполнение оценок (ученик - предметы)
   function FillMarksPp(Clas, ID: integer; sg: TAdvStringGrid): byte;
   // Заполнение оценок (предмет - ученики)
   function FillMarksPr(Clas, ID: integer; sg: TAdvStringGrid): byte;
   // Сохранение оценок все на все
   function MarksAllSave(sg: TAdvStringGrid): byte;
   // Сохранение оценок (0-по ученикам / 1-по предметам)
   function MarksAllSaveP(tp: integer; sg: TAdvStringGrid): byte;
   // Объединение классов
   function ClassO(F, S: integer): byte;
   // Удаление класса
   function ClassDel(ID: integer): byte;
   // Удаление ученика
   function PeopDel(ID: integer): byte;
   // Сохранение оценок в архив
   function MarksInArhiv(ID: integer): byte;
   // Заполнение табов
   function FillTabsText(Tab: TAdvOfficeTabSet; What, Table: string; Where: string=''; Order: string=''): byte;
   // Перевод на след. год
   function GoNextYear(): byte;
   // Мед-книга
   function FillMedAll(Clas: integer; sg: TAdvStringGrid; where: string = ''): byte;
   function FillMed(ID: integer; sg: TAdvStringGrid; where: string = ''): byte;
   // Заполнение прошлых классов для архива оценок
   function FillAClassCb(cb: TAdvOfficeComboBox; ID: integer): byte;
   // Заполнение архива оценок
   function FillAMarks(ID, Clas: integer; sg: TAdvStringGrid): byte;
   // Средняя успеваемость класса
   function SrClassMark(ID: integer): real;
   // Средняя успеваемость школы
   function SrSchoolMark(ID: integer): real;
   // Заполнение грида количества и успеваемости по паралелям
   function FillSGSchoolStat(sg: TAdvStringGrid; ID: integer): byte;
   // герб школы
   function FillImg(img: TImage; ID: integer): byte;
   // изменение изображения
   function cUpdateImg(ID: integer; Table: string; img: TImage): byte;
   // удаление школы
   function DeleteSchool(ID: integer): byte;
   // поиск учеников
   function Find(sg: TAdvStringGrid; What: string = ''): byte;
   // Изменение языковых настроек
   function ChangeLanguage(Path: string): byte;
   // Экспорт личного дела в Ворд
   function ExpWord(t: byte; ID: integer): byte;
   // Проверка вхождения класса в число прыгающих
   function CheckJump(cl: integer): byte;
   function SrClassAMark(ID: integer): real;
  end;

  Marks = record
   Peop: integer;
   Clas: integer;
   Predmet: string;
   Teacher: string;
   O1: string;
   O2: string;
   YER: string;
   FL: integer;
  end;
var
  fData: TfData;

implementation

uses BaseGrid, uMain;

{$R *.dfm}

{ Применение конфигурации }
function TfData.SetConfig: byte;
var f,e: TIniFile;
    ers: TStringList;
    i: integer;
begin
 try
  Result := 0;
  // открываем файл конфигурации
  f := TIniFile.Create(ExtractFileDir(Application.ExeName)+'\config.ini');
  // Читаем допустимую длину лог-файла
  LengthLog := f.ReadInteger('Settings', 'LengthLog', 10485760);
  // читаем текущий язык и запоминаем его для последующих действий
  language := f.ReadString('Language','Current','rus');
  // Максимальный номер класса
  MaxClass := f.ReadInteger('Settings', 'MaxClass', 12);
  // Забиваем массив по "перепрыгиванию" классов
  SetLength(JumpClass, MaxClass);
  for i := 0 to MaxClass-1 do
   JumpClass[i] := f.ReadInteger('JumpClass', IntToStr(i), 0);
  // читаем и сохраняем сообщения об ошибках
  e := TIniFile.Create(ExtractFileDir(Application.ExeName)+'\language\'+language+'.ini');
  SetLength(errors,0);
  ers := TStringList.Create;
  e.ReadSection('ERRORS',ers);
  SetLength(errors,ers.Count);
  for i := 0 to ers.Count-1 do
   errors[i] := e.ReadString('ERRORS',IntToStr(i),'ERROR');
  // читаем и сохраняем стандартные надписи
  ers.Clear;
  e.ReadSection('STANDART',ers);
  SetLength(standarts,ers.Count);
  for i := 0 to ers.Count-1 do
   standarts[i] := e.ReadString('STANDART',IntToStr(i),'ERROR');
  ers.Free;
  e.Free;
  // читаем настройки подключения к БД и подключаемся :)
  Database.Params.Clear;
  Term := 0;
  if not(FileExists(f.ReadString('DataBase','Path','base.fdb'))) then
  begin
   ShowError(2);
   Term := 1;
   Application.Terminate;
   Exit;
  end;
  Database.DatabaseName := f.ReadString('DataBase','Path','base.fdb');
  Database.Params.Add('user_name='+f.ReadString('DataBase','user_name','SYSDBA'));
  Database.Params.Add('password='+f.ReadString('DataBase','password','masterkey'));
  Database.Params.Add('lc_ctype='+f.ReadString('DataBase','lc_ctype','WIN1251'));
  try
   Database.Open();
  except
   Result := 1;
   ShowError(2);
  end;
  // чистим память
  f.Free;
  if (Result = 1) then begin
  Term := 1;
  Application.Terminate;
  end;
 except
  Result := 1;
  ShowError(1);
  SetLength(errors,0);
  SetLength(standarts,0);
  Term := 1;
   Application.Terminate;
   Exit;
 end;
end;

{ Установка языка для конкретной формы }
function TfData.SetLanguage(fm: TForm): byte;
var f: TIniFile;
    Params: TStringList;
    i: integer;
    tmp: TComponent;
    comp,prm: string;
begin
 try
  Result := 0;
  // открываем файл с языковыми настройками
  f := TIniFile.Create(ExtractFileDir(Application.ExeName)+'\language\'+language+'.ini');
  // читаем капшны компонентов
   Params := TStringList.Create;
   f.ReadSection(fm.Name,Params);
   // устанавливаем параметры компонентов
   for i := 0 to Params.Count-1 do
   begin
    // выделяем название компонента
    comp := Copy(Params[i],1,Pos('.',Params[i])-1);
    // выделяем изменяемый параметр
    prm := Copy(Params[i],Pos('.',Params[i])+1,Length(Params[i]));
    // ищем требуемый компонент на форме
    tmp := fm.FindComponent(comp);
    if (comp = 'TForm')                        then fm.Caption := f.ReadString(fm.Name,Params[i],'ERROR');
    // Если нашли компонент - сравниваем классы
    if (tmp <> nil) then
    begin
    if (tmp.ClassName = 'TAdvOfficePage')      then (tmp as TAdvOfficePage).Caption := f.ReadString(fm.Name,Params[i],'ERROR');
    if (tmp.ClassName = 'TAdvOfficeTabSet')    then (tmp as TAdvOfficeTabSet).AdvOfficeTabs[0].Caption := f.ReadString(fm.Name,Params[i],'ERROR');
    if (tmp.ClassName = 'TMenuItem')           then (tmp as TMenuItem).Caption := f.ReadString(fm.Name,Params[i],'ERROR');
    if (tmp.ClassName = 'TAdvReflectionLabel') then (tmp as TAdvReflectionLabel).HTMLText.Text := f.ReadString(fm.Name,Params[i],'ERROR');
    if (tmp.ClassName = 'TAdvGlassButton')     then (tmp as TAdvGlassButton).Caption := f.ReadString(fm.Name,Params[i],'ERROR');
    if (tmp.ClassName = 'TLabel')              then (tmp as TLabel).Caption := f.ReadString(fm.Name,Params[i],'ERROR');
    // в соответствии с параметром применяем значение
    if (tmp.ClassName = 'TAdvGlowButton') then
     case prm[1] of
      'C': (tmp as TAdvGlowButton).Caption := f.ReadString(fm.Name,Params[i],'ERROR');
      'H': (tmp as TAdvGlowButton).Hint    := f.ReadString(fm.Name,Params[i],'ERROR');
     end;
    if (tmp.ClassName = 'TAdvOfficeRadioGroup') then
     case prm[1] of
      'C': (tmp as TAdvOfficeRadioGroup).Caption  := f.ReadString(fm.Name,Params[i],'ERROR');
      '0': (tmp as TAdvOfficeRadioGroup).Items[0] := f.ReadString(fm.Name,Params[i],'ERROR');
      '1': (tmp as TAdvOfficeRadioGroup).Items[1] := f.ReadString(fm.Name,Params[i],'ERROR');
     end;
   end;
   end;
   // освобождаем
   Params.Free;
  // чистим память
  f.Free;
 except
  Result := 1;
  ShowError(3);
 end;
end;

procedure TfData.DataModuleCreate(Sender: TObject);
begin
 SetConfig();
end;

procedure TfData.ShowError(Error: integer; Handle: HWND);
begin
 MessageBox(Application.Handle, PChar(errors[Error]), PChar(errors[0]), MB_ICONWARNING or MB_OK);
end;

function TfData.FillSchollsTab(Tab: TAdvOfficeTabSet): byte;
begin
 try
  Result := 0;
  Tab.AdvOfficeTabs.Clear;
  try
   Tab.AdvOfficeTabs.Add();
   Tab.AdvOfficeTabs[0].Tag     := -1;
   Tab.AdvOfficeTabs[0].Caption := standarts[0];
   SQL.SQL.Add('Select * from TB_SCHOOL order by NAME');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    Tab.AdvOfficeTabs.Add();
    Tab.AdvOfficeTabs[Tab.AdvOfficeTabs.Count-1].Tag     := SQL.FieldByName('ID').AsInteger;
    Tab.AdvOfficeTabs[Tab.AdvOfficeTabs.Count-1].Caption := SQL.FieldByName('NAME').AsString;
    SQL.Next();
   end;
   if (Tab.AdvOfficeTabs.Count > 1) then Tab.AdvOfficeTabs.Items[0].Free else Result := 1;
   Tab.ActiveTabIndex := 0;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear;
  SQL.Close;
  Transaction.Active := FALSE;
 end;
end;

function TfData.FillClassTab(School: integer; Tab: TAdvOfficeTabSet; Where: string): byte;
begin
 try
  Result := 0;
  Tab.AdvOfficeTabs.Clear;
  try
   Tab.AdvOfficeTabs.Add();
   Tab.AdvOfficeTabs[0].Tag     := -1;
   Tab.AdvOfficeTabs[0].Caption := standarts[0];
   SQL.SQL.Add('Select ID,NUM,NAME from TB_CLASS where (SCHOOL='+IntToStr(School)+') '+Trim(Where)+' order by NUM,NAME');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    Tab.AdvOfficeTabs.Add();
    Tab.AdvOfficeTabs[Tab.AdvOfficeTabs.Count-1].Tag     := SQL.FieldByName('ID').AsInteger;
    Tab.AdvOfficeTabs[Tab.AdvOfficeTabs.Count-1].Caption := SQL.FieldByName('NUM').AsString+'-'+SQL.FieldByName('NAME').AsString;
    SQL.Next();
   end;
   if (Tab.AdvOfficeTabs.Count > 1) then Tab.AdvOfficeTabs.Items[0].Free else Result := 1;
   Tab.ActiveTabIndex := 0;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.FillFIOSg(Table: string; Filter: string; ID: integer; sg: TAdvStringGrid; andWhere: string; cwidth: integer): byte;
begin
 try
  Result := 0;
  sg.ColWidths[0] := cwidth;
  sg.RowCount := 1;
  sg.ColCount := 1;
  try
   sg.Objects[0,0] := Pointer(-1);
   sg.Cells[0,0] := '<P align="center">'+standarts[1]+'</p>';
   SQL.SQL.Add('Select ID,FNAME,NAME,SNAME from '+Trim(Table)+' where ('+Trim(Filter)+'='+IntToStr(ID)+') '+Trim(andWhere)+' order by FNAME,NAME,SNAME');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    sg.RowCount := sg.RowCount+1;
    sg.Objects[0,sg.RowCount-1] := Pointer(SQL.FieldByName('ID').AsInteger);
    sg.Cells[0,sg.RowCount-1] := SQL.FieldByName('FNAME').AsString+' '+SQL.FieldByName('NAME').AsString[1]+'. '+SQL.FieldByName('SNAME').AsString[1]+'.';
    SQL.Next();
   end;
   if (sg.RowCount > 1) then sg.FixedRows := 1 else Result := 1;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.FillSchollsSG(SG: TAdvStringGrid): byte;
var i: integer;
begin
 try
  Result := 0;
  SG.RowCount := 1;  i := 1;
  SG.Cells[0,0] := standarts[0]; SG.Objects[0,0] := Pointer(-1);
  try
   SQL.SQL.Add('Select * from TB_SCHOOL order by NAME');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    SG.RowCount := i;
    SG.Objects[0,i-1] := Pointer(SQL.FieldByName('ID').AsInteger);
    SG.Cells[0,i-1]   := '     '+SQL.FieldByName('NAME').AsString;
    SQL.Next();
    Inc(i);
   end;
   if ((SG.RowCount = 1)and(SG.Cells[0,0] = standarts[0])) then Result := 1;
   SG.Row := 0;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear;
  SQL.Close;
  Transaction.Active := FALSE;
 end;
end;

function TfData.cSelectS(What, Table, Where: string): string;
begin
 try
  Result := 'OK';
  try
   SQL.SQL.Add('Select '+Trim(What)+' as RS from '+Trim(Table)+' '+Trim(Where)+';');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   Result := SQL.FieldByName('RS').AsString;
  except
   Result := 'ERROR';
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear;
  SQL.Close;
  Transaction.Active := FALSE;
 end;
end;

function TfData.cFillCb(cb: TAdvOfficeComboBox; What, Table, Where: string): byte;
begin
 try
  Result := 0;
  cb.Clear;
  cb.Items.Clear;
  cb.Items.AddObject(standarts[0], Pointer(-1));
  try
   SQL.SQL.Add('Select ID, '+Trim(What)+' from '+Trim(Table)+' '+Trim(Where)+' order by '+Trim(What)+';');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    cb.Items.AddObject(SQL.FieldByName(What).AsString, Pointer(SQL.FieldByName('ID').AsInteger));
    SQL.Next();
   end;
   if (cb.Items.Count > 1) then cb.Items.Delete(0) else Result := 1;
   cb.ItemIndex := 0;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.FillClassCb(cb: TAdvOfficeComboBox; Where: string): byte;
begin
 try
  Result := 0;
  cb.Clear;
  cb.Items.Clear;
  cb.Items.AddObject(standarts[0], Pointer(-1));
  try
   SQL.SQL.Clear();
   SQL.Close();
   Transaction.Active := FALSE;
   SQL.SQL.Add('Select ID,NUM,NAME from TB_CLASS '+Trim(Where)+' order by NUM,NAME;');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    cb.Items.AddObject(SQL.FieldByName('NUM').AsString+'-'+SQL.FieldByName('NAME').AsString, Pointer(SQL.FieldByName('ID').AsInteger));
    SQL.Next();
   end;
   SQL.SQL.Clear();
   SQL.Close();
   Transaction.Active := FALSE;
   if (cb.Items.Count > 1) then cb.Items.Delete(0) else Result := 1;
   cb.ItemIndex := 0;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.cInsert(Table, Rec, Val: string): byte;
begin
 try
  Result := 0;
  try
   SQL.SQL.Add('Insert into '+Trim(Table)+' ('+Trim(Rec)+') values ('+Trim(Val)+');');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   Transaction.Commit();
  except
   Result := 1;
   Transaction.Rollback();
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear;
  SQL.Close;
  Transaction.Active := FALSE;
 end;
end;

function TfData.cInserts(Table: string; Recs, Vals: TStringList): byte;
var i: integer;
    r,v: string;
begin
 try
  Result := 0;
  r := '';  v := '';
  try
   for i := 0 to Recs.Count-1 do r := r+Recs[i]+','; delete(r,Length(r),1);
   for i := 0 to Vals.Count-1 do v := v+Vals[i]+','; delete(v,Length(v),1);
   SQL.SQL.Add('Insert into '+Trim(Table)+' ('+Trim(r)+') values ('+Trim(v)+');');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   Transaction.Commit();
  except
   Result := 1;
   Transaction.Rollback();
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear;
  SQL.Close;
  Transaction.Active := FALSE;
 end;
end;

function TfData.cUpdates(Table: string; Recs, Vals: TStringList; Where: string): byte;
var i: integer;
    s: string;
begin
 try
  Result := 0;
  s := '';
  try
   for i := 0 to Recs.Count-1 do s := s+Recs[i]+'='+Vals[i]+','; delete(s,Length(s),1);
   SQL.SQL.Add('Update '+Trim(Table)+' set '+Trim(s)+' '+Trim(Where)+';');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   Transaction.Commit();
  except
   Result := 1;
   Transaction.Rollback();
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear;
  SQL.Close;
  Transaction.Active := FALSE;
 end;
end;

function TfData.FillPlanSg(ID: integer; sg: TAdvStringGrid; Order: string): byte;
begin
 try
  Result := 0;
  sg.RowCount := 1;
  sg.ColCount := 3;
  // Айдишник
  sg.ColWidths[0] := -1;
  // Учитель
  sg.ColWidths[1] := 150;
  // Предмет
  sg.ColWidths[2] := 150;
  try
   sg.Objects[0,0] := Pointer(-1);
   sg.Cells[1,0] := '<P align="center">'+standarts[7]+'</p>';
   sg.Cells[2,0] := '<P align="center">'+standarts[8]+'</p>';
   SQL.SQL.Add('Select distinct TB_PLANCLASS.ID AS UID,TB_TEACHER.ID AS TID,TB_TEACHER.FNAME AS TF,TB_TEACHER.NAME AS TN,TB_TEACHER.SNAME AS TS,TB_PREDMET.ID AS PID,TB_PREDMET.NAME AS PN from TB_PLANCLASS,TB_UPLAN,TB_TEACHER,TB_PREDMET where ((TB_PLANCLASS.UPLAN='+IntToStr(ID)+')and(TB_PLANCLASS.TEACHER=TB_TEACHER.ID)and(TB_PREDMET.ID=TB_PLANCLASS.PREDMET)) order by TB_TEACHER.FNAME, TB_TEACHER.NAME, TB_TEACHER.SNAME');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    sg.RowCount := sg.RowCount+1;
    // Айдишник
    sg.Objects[0,sg.RowCount-1] := Pointer(SQL.FieldByName('UID').AsInteger);
    // Учитель Ф.И.О.
    sg.Objects[1,sg.RowCount-1] := Pointer(SQL.FieldByName('TID').AsInteger);
    sg.Cells[1,sg.RowCount-1] := SQL.FieldByName('TF').AsString+' '+SQL.FieldByName('TN').AsString[1]+'. '+SQL.FieldByName('TS').AsString[1]+'.';
    // Название предмета
    sg.Objects[2,sg.RowCount-1] := Pointer(SQL.FieldByName('PID').AsInteger);
    sg.Cells[2,sg.RowCount-1] := SQL.FieldByName('PN').AsString;
    SQL.Next();
   end;
   if (sg.RowCount > 1) then sg.FixedRows := 1 else Result := 1;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.FillCbPredmetPlan(cb: TAdvOfficeComboBox; ID: integer): byte;
begin
 try
  Result := 0;
  cb.Clear;
  cb.Items.Clear;
  cb.Items.AddObject(standarts[0], Pointer(-1));
  try
   SQL.SQL.Add('Select TB_PREDMET.ID,TB_PREDMET.NAME from TB_PREDMET where (ID not in (Select TB_PLANCLASS.PREDMET from TB_PLANCLASS where TB_PLANCLASS.UPLAN='+IntToStr(ID)+')) order by NAME');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    cb.Items.AddObject(SQL.FieldByName('NAME').AsString, Pointer(SQL.FieldByName('ID').AsInteger));
    SQL.Next();
   end;
   if (cb.Items.Count > 1) then cb.Items.Delete(0) else Result := 1;
   cb.ItemIndex := 0;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.FillFIOCb(Table, Filter: string; ID: integer; cb: TAdvOfficeComboBox): byte;
begin
 try
  Result := 0;
  cb.Clear;
  cb.Items.Clear;
  cb.Items.AddObject(standarts[0], Pointer(-1));
  try
   SQL.SQL.Add('Select ID,FNAME,NAME,SNAME from '+Trim(Table)+' where '+Trim(Filter)+'='+IntToStr(ID)+' order by FNAME,NAME,SNAME');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    cb.Items.AddObject(SQL.FieldByName('FNAME').AsString+' '+SQL.FieldByName('NAME').AsString[1]+'. '+SQL.FieldByName('SNAME').AsString[1]+'.', Pointer(SQL.FieldByName('ID').AsInteger));
    SQL.Next();
   end;
   if (cb.Items.Count > 1) then cb.Items.Delete(0) else Result := 1;
   cb.ItemIndex := 0;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.cFillSg(sg: TAdvStringGrid; Er: integer; What, Table, Where: string): byte;
var i: integer;
begin
 try
  Result := 0;
  SG.RowCount := 2;  i := 2;
  SG.Cells[0,0] := standarts[Er]; SG.Objects[0,0] := Pointer(-1);
  SG.Cells[0,1] := standarts[0];  SG.Objects[0,1] := Pointer(-1);
  try
   SQL.SQL.Add('Select ID, '+Trim(What)+' AS RS from '+Trim(Table)+' '+Trim(Where)+' order by '+Trim(What)+';');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    SG.RowCount := i;
    SG.Objects[0,i-1] := Pointer(SQL.FieldByName('ID').AsInteger);
    SG.Cells[0,i-1]   := '     '+SQL.FieldByName('RS').AsString;
    SQL.Next();
    Inc(i);
   end;
   if ((SG.RowCount = 2)and(integer(SG.Objects[0,sg.RowCount-1]) = -1)) then Result := 1;
   SQL.SQL.Clear;
   SQL.Close;
   Transaction.Active := FALSE;
   SG.Row := 0;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear;
  SQL.Close;
  Transaction.Active := FALSE;
 end;
end;

function TfData.cUpdate(Table, Rec, Val, Where: string): byte;
begin
 try
  Result := 0;
  try
   SQL.SQL.Add('Update '+Trim(Table)+' set '+Trim(Rec)+'='+Trim(Val)+' '+Trim(Where)+';');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   Transaction.Commit();
  except
   Result := 1;
   Transaction.Rollback();
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear;
  SQL.Close;
  Transaction.Active := FALSE;
 end;
end;

function TfData.cDelete(Table, Where: string): byte;
begin
 try
  Result := 0;
  try
   SQL.SQL.Add('Delete from '+Trim(Table)+' '+Trim(Where)+';');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   Transaction.Commit();
  except
   Result := 1;
   Transaction.Rollback();
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear;
  SQL.Close;
  Transaction.Active := FALSE;
 end;
end;

function TfData.cCount(What, Table, Where: string): integer;
begin
 try
  Result := 0;
  try
   SQL.SQL.Add('Select Count('+Trim(What)+') AS RS from '+Trim(Table)+' '+Trim(Where)+' ;');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   Result := SQL.FieldByName('RS').AsInteger;
  except
   Result := -1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear;
  SQL.Close;
  Transaction.Active := FALSE;
 end;
end;

function TfData.FillPlanSGf(tp: byte; School: Integer; ID: Integer; sg: TAdvStringGrid): byte;
begin
 try
  Result := 0;
  sg.RowCount := 1;
  sg.ColCount := 3;
  // Айдишник
  sg.ColWidths[0] := -1;
  // Учитель
  sg.ColWidths[1] := 150;
  // Предмет
  sg.ColWidths[2] := 150;
  try
   sg.Objects[0,0] := Pointer(-1);
   case tp of
    0: begin
        sg.Cells[1,0] := '<P align="center">'+standarts[12]+'</p>';
        sg.Cells[2,0] := '<P align="center">'+standarts[7]+'</p>';
       end;
    1: begin
        sg.Cells[1,0] := '<P align="center">'+standarts[12]+'</p>';
        sg.Cells[2,0] := '<P align="center">'+standarts[8]+'</p>';
       end;
   end;
   
   case tp of
    0: begin
        SQL.SQL.Add('Select TB_PLANCLASS.ID AS UID,TB_UPLAN.ID AS IDf,TB_PLANCLASS.TEACHER AS IDs,TB_CLASS.NUM AS CNU,TB_CLASS.NAME AS CNM,TB_TEACHER.FNAME AS TF,TB_TEACHER.NAME AS TN,TB_TEACHER.SNAME AS TS from ');
        SQL.SQL.Add('TB_UPLAN,TB_PLANCLASS,TB_CLASS,TB_TEACHER where((TB_PLANCLASS.PREDMET='+IntToStr(ID)+')and(TB_PLANCLASS.UPLAN=TB_UPLAN.ID)and(TB_UPLAN.ID=TB_CLASS.UPLAN)and(TB_PLANCLASS.TEACHER=TB_TEACHER.ID)and(TB_CLASS.SCHOOL='+IntToStr(School)+')and(TB_TEACHER.SCHOOL='+IntToStr(School)+')and(TB_CLASS.NUM<'+IntToStr(MaxClass)+')) order by TB_CLASS.NUM, TB_CLASS.NAME;');
       end;
    1: SQL.SQL.Add('Select TB_PLANCLASS.ID AS UID, TB_UPLAN.ID AS IDf, TB_PLANCLASS.PREDMET AS IDs, TB_CLASS.NUM AS CNU, TB_CLASS.NAME AS CNM, TB_PREDMET.NAME AS PN from TB_UPLAN, TB_CLASS, TB_PLANCLASS, TB_PREDMET where((TB_PLANCLASS.TEACHER='+IntToStr(ID)+')and(TB_UPLAN.ID=TB_PLANCLASS.UPLAN)and(TB_UPLAN.ID=TB_CLASS.UPLAN)and(TB_PLANCLASS.PREDMET=TB_PREDMET.ID)and(TB_CLASS.SCHOOL='+IntToStr(School)+')and(TB_CLASS.NUM<'+IntToStr(MaxClass)+')) order by TB_CLASS.NUM, TB_CLASS.NAME;');
   end;

   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while (not SQL.Eof) do
   begin
    sg.RowCount := sg.RowCount+1;
    // Айдишник
    sg.Objects[0,sg.RowCount-1] := Pointer(SQL.FieldByName('UID').AsInteger);
    // Класс
    sg.Objects[1,sg.RowCount-1] := Pointer(SQL.FieldByName('IDf').AsInteger);
    sg.Cells[1,sg.RowCount-1] := SQL.FieldByName('CNU').AsString+'-'+SQL.FieldByName('CNM').AsString;
    // Фильт
    sg.Objects[2,sg.RowCount-1] := Pointer(SQL.FieldByName('IDs').AsInteger);
    case tp of
     0: // Учитель
        sg.Cells[2,sg.RowCount-1] := SQL.FieldByName('TF').AsString+' '+SQL.FieldByName('TN').AsString[1]+'. '+SQL.FieldByName('TS').AsString[1]+'.';
     1: // Предмет
        sg.Cells[2,sg.RowCount-1] := SQL.FieldByName('PN').AsString;
    end;
    SQL.Next();
   end;
   if (sg.RowCount > 1) then sg.FixedRows := 1 else Result := 1;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.ExpPlansExcel(tp: byte; ID: integer; sg: TAdvStringGrid): byte;
var ExcelApp, Workbook: Variant;
    i: integer;
    tmp: string;
begin
 try
  Result := 0;
  // Создание Excel
  ExcelApp := CreateOleObject('Excel.Application');
  // Отключаем реакцию Excel на события, чтобы ускорить вывод информации
  ExcelApp.Application.EnableEvents := FALSE;
  ExcelApp.Application.DisplayAlerts := FALSE;
  //  Создаем Книгу (Workbook)
  Workbook := ExcelApp.WorkBooks.Add;
  // Делаем Excel невидимым
  ExcelApp.Visible := FALSE;
  case tp of
   0: ExcelApp.WorkBooks[1].WorkSheets[1].Name := cSelectS('NAME','TB_UPLAN','where ID='+IntToStr(ID));
   1: ExcelApp.WorkBooks[1].WorkSheets[1].Name := cSelectS('NAME','TB_PREDMET','where ID='+IntToStr(ID)+' ');
   2: ExcelApp.WorkBooks[1].WorkSheets[1].Name := cSelectS('FNAME','TB_TEACHER','where ID='+IntToStr(ID)+' ')+' '+cSelectS('NAME','TB_TEACHER','where ID='+IntToStr(ID)+' ')+' '+cSelectS('SNAME','TB_TEACHER','where ID='+IntToStr(ID)+' ');
  end;
  // Формируем шапку
  ExcelApp.Range['A1:B1'].HorizontalAlignment := 3;
  ExcelApp.Range['A1:B1'].Borders.LineStyle := 1;
  ExcelApp.ActiveWorkBook.WorkSheets[1].Columns[1].ColumnWidth := 25;
  ExcelApp.ActiveWorkBook.WorkSheets[1].Columns[2].ColumnWidth := 25;
  tmp := sg.Cells[1, 0];
  delete(tmp,1,18);
  delete(tmp,pos('<',tmp),Length(tmp));
  ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, 1]  := tmp;
  tmp := sg.Cells[2, 0];
  delete(tmp,1,18);
  delete(tmp,pos('<',tmp),Length(tmp));
  ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, 2]  := tmp;
  // Копируем инфу из грида в лист
  for i := 1 to sg.RowCount-1 do
  begin
   ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[i+1, 1]  := sg.Cells[1, i];
   ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[i+1, 2]  := sg.Cells[2, i];
  end;
  ExcelApp.Range['A1:B'+IntToStr(i+1)].WrapText := True;
  // Делаем Excel видимым
  ExcelApp.Application.EnableEvents := TRUE;
  ExcelApp.Application.DisplayAlerts := TRUE;
  ExcelApp.Visible := TRUE;
 except
  Result := 1;
  ShowError(4);
 end;
end;

function TfData.FillFIOMarksSG(ID: integer; sg: TAdvStringGrid; sgw: integer): byte;
begin
 try
  Result := 0;
  sg.ColWidths[0] := 110;
  sg.RowCount := 2;
  sg.ColCount := 1;
  try
   sg.Objects[0,0] := Pointer(-1);
   sg.Cells[0,0] := '<P align="center">'+standarts[1]+'</p>';
   sg.Objects[0,1] := Pointer(-1);
   sg.Cells[0,1] := standarts[13];
   sg.FixedRows := 1;
   SQL.SQL.Add('Select ID,FNAME,NAME,SNAME from TB_PEOP where (CLASS='+IntToStr(ID)+') order by FNAME,NAME,SNAME');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    sg.RowCount := sg.RowCount+1;
    sg.Objects[0,sg.RowCount-1] := Pointer(SQL.FieldByName('ID').AsInteger);
    sg.Cells[0,sg.RowCount-1] := SQL.FieldByName('FNAME').AsString+' '+SQL.FieldByName('NAME').AsString[1]+'. '+SQL.FieldByName('SNAME').AsString[1]+'.';
    SQL.Next();
   end;
   if (sg.RowCount = 2) then Result := 1;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.FillPrMarksSG(ID: integer; sg: TAdvStringGrid): byte;
begin
 try
  Result := 0;
  sg.ColWidths[0] := 110;
  sg.RowCount := 2;
  sg.ColCount := 1;
  try
   sg.Objects[0,0] := Pointer(-1);
   sg.Cells[0,0] := '<P align="center">'+standarts[1]+'</p>';
   sg.Objects[0,1] := Pointer(-1);
   sg.Cells[0,1] := standarts[13];
   sg.FixedRows := 1;
   SQL.SQL.Add('Select TB_PREDMET.ID AS PID, TB_PREDMET.NAME AS PN from TB_PREDMET,TB_CLASS,TB_PLANCLASS,TB_UPLAN where ((TB_CLASS.ID='+IntToStr(ID)+')and(TB_UPLAN.ID=TB_CLASS.UPLAN)and(TB_PLANCLASS.UPLAN=TB_UPLAN.ID)and(TB_PLANCLASS.PREDMET=TB_PREDMET.ID)) order by TB_PREDMET.NAME');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    sg.RowCount := sg.RowCount+1;
    sg.Objects[0,sg.RowCount-1] := Pointer(SQL.FieldByName('PID').AsInteger);
    sg.Cells[0,sg.RowCount-1] := SQL.FieldByName('PN').AsString;
    SQL.Next();
   end;
   if (sg.RowCount = 2) then Result := 1;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.FillMarksAll(ID: integer; sg: TAdvStringGrid): byte;
var i, j: integer;
begin
 try
  Result := 0;
  sg.SplitAllCells();
  sg.RowCount := 3;
  sg.ColCount := 2;
  sg.ColWidths[1] := 150;
  try
   // Формируем шапку
   { ------------|-----------------------------|-------
                 |    Название предмета        |
                 |-----------------------------|-------
                 |  1семестр | 2 семестр | Год |
     --------------------------------------------------
     Ученик      |    10     |    11     |  11 |
     ---------------------------------------------------
   }
   sg.MergeCells(1,1,1,2);
   sg.Objects[0,0] := Pointer(ID);
   sg.Cells[1,1] := '<P align="center">'+standarts[14]+'</p>';
//   sg.FixedRows := 2;
   sg.FixedCols := 1;
   sg.RowHeights[0] := -1;
   sg.ColWidths[0] := -1;
   
   // Выбираем учеников
   SQL.SQL.Add('Select ID,FNAME,NAME,SNAME from TB_PEOP where (CLASS='+IntToStr(ID)+') order by FNAME,NAME,SNAME');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    sg.RowCount := sg.RowCount+1;
    sg.Objects[1,sg.RowCount-1] := Pointer(SQL.FieldByName('ID').AsInteger);
    sg.Cells[1,sg.RowCount-1] := SQL.FieldByName('FNAME').AsString+' '+SQL.FieldByName('NAME').AsString[1]+'. '+SQL.FieldByName('SNAME').AsString[1]+'.';
    SQL.Next();
   end;
   SQL.SQL.Clear();
   SQL.Close();
   Transaction.Active := FALSE;

   // Выбираем предметы
   SQL.SQL.Add('Select TB_PREDMET.ID AS PID, TB_PREDMET.NAME AS PN,TB_TEACHER.FNAME AS TFN, TB_TEACHER.NAME AS TN, TB_TEACHER.SNAME AS TSN from TB_PREDMET, TB_CLASS, TB_PLANCLASS ,TB_UPLAN,TB_TEACHER where ((TB_CLASS.ID='+IntToStr(ID)+')and(TB_UPLAN.ID=TB_CLASS.UPLAN)and(TB_PLANCLASS.UPLAN=TB_UPLAN.ID)and(TB_PLANCLASS.PREDMET=TB_PREDMET.ID)and(TB_TEACHER.ID=TB_PLANCLASS.TEACHER)) order by TB_PREDMET.NAME');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    sg.ColCount := sg.ColCount+3;
    sg.MergeCells(sg.ColCount-3,1,3,1);
    sg.Objects[sg.ColCount-1, 1] := Pointer(SQL.FieldByName('PID').AsInteger);
    sg.RowHeights[1] := 44;
    sg.Cells[sg.ColCount-1, 1] := '<P align="center">'+SQL.FieldByName('PN').AsString+'</P>'+#10#13+'<P align="center">'+SQL.FieldByName('TFN').AsString+' '+SQL.FieldByName('TN').AsString[1]+'. '+SQL.FieldByName('TSN').AsString+'.'+'</p>';
    sg.Cells[sg.ColCount-3, 2] := '<P align="center">'+standarts[15]+'</p>';
    sg.Cells[sg.ColCount-2, 2] := '<P align="center">'+standarts[16]+'</p>';
    sg.Cells[sg.ColCount-1, 2] := '<P align="center">'+standarts[17]+'</p>';
    SQL.Next();
   end;
   SQL.SQL.Clear();
   SQL.Close();
   Transaction.Active := FALSE;

   // Выбираем оценки
   // Бежим по ученикам вниз
   for i := 3 to sg.RowCount-1 do
   begin
    j := 5;
    // Бежим по предметам, перескакивая по 3 ячейки
    while (j <= sg.ColCount) do
    begin
     SQL.SQL.Add('Select ID, O1, O2, YER from TB_MARKS where ((PEOP='+IntToStr(integer(sg.Objects[1, i]))+')and(PREDMET='+IntToStr(integer(sg.Objects[j-1, 1]))+'))');
     Transaction.Active := TRUE;
     SQL.ExecQuery();
     sg.Objects[j-1, i] := Pointer(SQL.FieldByName('ID').AsInteger);
     if (SQL.FieldByName('O1').AsString = '') then sg.Cells[j-3, i] := '0' else
     sg.Cells[j-3, i]   := SQL.FieldByName('O1').AsString;
     if (SQL.FieldByName('O2').AsString = '') then sg.Cells[j-2, i] := '0' else
     sg.Cells[j-2, i]   := SQL.FieldByName('O2').AsString;
     if (SQL.FieldByName('YER').AsString = '') then sg.Cells[j-1, i] := '0' else
     sg.Cells[j-1, i]   := SQL.FieldByName('YER').AsString;
     SQL.SQL.Clear();
     SQL.Close();
     Transaction.Active := FALSE;
     j := j + 3;
    end;
   end;

   if (sg.RowCount > 3) then sg.FixedRows := 3 else sg.FixedRows := 2;
   if (sg.ColCount > 2) then sg.FixedRows := 2 else sg.FixedRows := 1;
   if ((sg.RowCount = 3)or(sg.ColCount = 1)) then Result := 1;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.MarksAllSave(sg: TAdvStringGrid): byte;
var i, j, k: integer;
    r, v: TStringList;
    t: string;
begin
 try
  Result := 0;
  r := TStringList.Create(); r.Clear();
  r.Add('PEOP'); r.Add('PREDMET'); r.Add('O1'); r.Add('O2'); r.Add('YER'); r.Add('FL');
  v := TStringList.Create(); v.Clear();
  try
   for i := 3 to sg.RowCount-1 do
   begin
    j := 5;
    // Бежим по предметам, перескакивая по 3 ячейки
    while (j <= sg.ColCount) do
    begin
     v.Clear();
     v.Add(IntToStr(integer(sg.Objects[1, i])));
     v.Add(IntToStr(integer(sg.Objects[j-1, 1])));
     // проверка
     t := sg.Cells[j-3, i];
     k := pos(',', t);
     if (k <> 0) then
     begin
      delete(t, k-1, 1);
      t := copy(t, 1, k-1) + '.' + copy(t, k+1, Length(t));
     end;
     sg.Cells[j-3, i] := t;

     t := sg.Cells[j-2, i];
     k := pos(',', t);
     if (k <> 0) then
     begin
      delete(t, k-1, 1);
      t := copy(t, 1, k-1) + '.' + copy(t, k+1, Length(t));
     end;
     sg.Cells[j-2, i] := t;

     t := sg.Cells[j-1, i];
     k := pos(',', t);
     if (k <> 0) then
     begin
      delete(t, k-1, 1);
      t := copy(t, 1, k-1) + '.' + copy(t, k+1, Length(t));
     end;
     sg.Cells[j-1, i] := t;
     // конец проверки

     v.Add(''''+sg.Cells[j-3, i]+'''');
     v.Add(''''+sg.Cells[j-2, i]+'''');
     v.Add(''''+sg.Cells[j-1, i]+'''');
     if ((pos('.',sg.Cells[j-3, i]) = 0)and(pos('.',sg.Cells[j-2, i]) = 0)and(pos('.',sg.Cells[j-1, i]) = 0)) then v.Add('0') else v.Add('1');

     if ((sg.Objects[j-1, i] = nil)or(integer(sg.Objects[j-1, i]) = 0)) then
     begin
     if (cInserts('TB_MARKS',r,v) = 1) then Exception.Create('');
     end else
     begin
      if (cUpdates('TB_MARKS',r,v,'where ID='+IntToStr(integer(sg.Objects[j-1, i]))) = 1) then Exception.Create('');
     end;
     j := j + 3;
    end;
   end;

  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
  r.Free; v.Free;
 end;
end;

function TfData.FillMarksPp(Clas: integer; ID: integer; sg: TAdvStringGrid): byte;
var i: integer;
begin
 try
  Result := 0;
  sg.SplitAllCells();
//  sg.
  sg.RowCount := 2;
  sg.ColCount := 5;
  sg.ColWidths[1] := 150;
  try
   // Формируем шапку
   { ----------------|-----------------------------|-------
                     |  1семестр | 2 семестр | Год |
     ---------------------------------------------------
     Ученик|Предмет  |    10     |    11     |  11 |
     ---------------------------------------------------
   }
   sg.Objects[0,0] := Pointer(ID);
   sg.Cells[1,1] := '<P align="center">'+standarts[18]+'</p>';
   // 1 сем., 2 сем., Год
   sg.Cells[2, 1] := '<P align="center">'+standarts[15]+'</p>';
   sg.Cells[3, 1] := '<P align="center">'+standarts[16]+'</p>';
   sg.Cells[4, 1] := '<P align="center">'+standarts[17]+'</p>';
//   sg.FixedRows := 2;
   sg.FixedCols := 1;
   sg.RowHeights[0] := -1;
   sg.ColWidths[0] := -1;

   // Выбираем предметы
   SQL.SQL.Add('Select TB_PREDMET.ID AS PID, TB_PREDMET.NAME AS PN from TB_PREDMET,TB_CLASS,TB_PLANCLASS,TB_UPLAN where ((TB_CLASS.ID='+IntToStr(Clas)+')and(TB_CLASS.UPLAN=TB_UPLAN.ID)and(TB_PLANCLASS.UPLAN=TB_UPLAN.ID)and(TB_PLANCLASS.PREDMET=TB_PREDMET.ID)) order by TB_PREDMET.NAME');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    sg.RowCount := sg.RowCount+1;
    sg.Objects[1, sg.RowCount-1] := Pointer(SQL.FieldByName('PID').AsInteger);
    sg.Cells[1, sg.RowCount-1] := SQL.FieldByName('PN').AsString;
    SQL.Next();
   end;
   SQL.SQL.Clear();
   SQL.Close();
   Transaction.Active := FALSE;

   // Выбираем оценки
   // Бежим по ученикам вниз
   for i := 2 to sg.RowCount-1 do
   begin
    SQL.SQL.Add('Select ID, O1, O2, YER from TB_MARKS where ((PEOP='+IntToStr(ID)+')and(PREDMET='+IntToStr(integer(sg.Objects[1, i]))+'))');
    Transaction.Active := TRUE;
    SQL.ExecQuery();
    sg.Objects[2, i] := Pointer(SQL.FieldByName('ID').AsInteger);
    if (SQL.FieldByName('O1').AsString = '') then sg.Cells[2, i] := '0' else
    sg.Cells[2, i]   := SQL.FieldByName('O1').AsString;
    if (SQL.FieldByName('O2').AsString = '') then sg.Cells[3, i] := '0' else
    sg.Cells[3, i]   := SQL.FieldByName('O2').AsString;
    if (SQL.FieldByName('YER').AsString = '') then sg.Cells[4, i] := '0' else
    sg.Cells[4, i]   := SQL.FieldByName('YER').AsString;
    SQL.SQL.Clear();
    SQL.Close();
    Transaction.Active := FALSE;
   end;

   if (sg.RowCount > 2) then sg.FixedRows := 2 else sg.FixedRows := 1;
   sg.FixedCols := 2;
   if (sg.RowCount = 2) then Result := 1;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.FillMarksPr(Clas: integer; ID: integer; sg: TAdvStringGrid): byte;
var i: integer;
begin
 try
  Result := 0;
  sg.SplitAllCells();
//  sg.
  sg.RowCount := 2;
  sg.ColCount := 5;
  sg.ColWidths[1] := 150;
  try
   // Формируем шапку
   { ----------------|-----------------------------|-------
                     |  1семестр | 2 семестр | Год |
     ---------------------------------------------------
     Ученик|Предмет  |    10     |    11     |  11 |
     ---------------------------------------------------
   }
   sg.Objects[0,0] := Pointer(ID);
   sg.Objects[1,1] := Pointer(Clas);
   sg.Cells[1,1] := '<P align="center">'+standarts[19]+'</p>';
   // 1 сем., 2 сем., Год
   sg.Cells[2, 1] := '<P align="center">'+standarts[15]+'</p>';
   sg.Cells[3, 1] := '<P align="center">'+standarts[16]+'</p>';
   sg.Cells[4, 1] := '<P align="center">'+standarts[17]+'</p>';
//   sg.FixedRows := 2;
   sg.FixedCols := 1;
   sg.RowHeights[0] := -1;
   sg.ColWidths[0] := -1;

   // Выбираем учеников
   SQL.SQL.Add('Select ID,FNAME,NAME,SNAME from TB_PEOP where (CLASS='+IntToStr(Clas)+') order by FNAME,NAME,SNAME');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    sg.RowCount := sg.RowCount+1;
    sg.Objects[1,sg.RowCount-1] := Pointer(SQL.FieldByName('ID').AsInteger);
    sg.Cells[1,sg.RowCount-1] := SQL.FieldByName('FNAME').AsString+' '+SQL.FieldByName('NAME').AsString[1]+'. '+SQL.FieldByName('SNAME').AsString[1]+'.';
    SQL.Next();
   end;
   SQL.SQL.Clear();
   SQL.Close();
   Transaction.Active := FALSE;

   // Выбираем оценки
   // Бежим по ученикам вниз
   for i := 2 to sg.RowCount-1 do
   begin
    SQL.SQL.Add('Select ID, O1, O2, YER from TB_MARKS where ((PEOP='+IntToStr(integer(sg.Objects[1, i]))+')and(PREDMET='+IntToStr(ID)+'))');
    Transaction.Active := TRUE;
    SQL.ExecQuery();
    sg.Objects[2, i] := Pointer(SQL.FieldByName('ID').AsInteger);
    if (SQL.FieldByName('O1').AsString = '') then sg.Cells[2, i] := '0' else
    sg.Cells[2, i]   := SQL.FieldByName('O1').AsString;
    if (SQL.FieldByName('O2').AsString = '') then sg.Cells[3, i] := '0' else
    sg.Cells[3, i]   := SQL.FieldByName('O2').AsString;
    if (SQL.FieldByName('YER').AsString = '') then sg.Cells[4, i] := '0' else
    sg.Cells[4, i]   := SQL.FieldByName('YER').AsString;
    SQL.SQL.Clear();
    SQL.Close();
    Transaction.Active := FALSE;
   end;

   if (sg.RowCount > 2) then sg.FixedRows := 2 else sg.FixedRows := 1;
   sg.FixedCols := 2;
   if (sg.RowCount = 2) then Result := 1;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.MarksAllSaveP(tp: integer; sg: TAdvStringGrid): byte;
var i, k: integer;
    r, v: TStringList;
    t: string;
begin
 try
  Result := 0;
  r := TStringList.Create(); r.Clear();
  r.Add('PEOP'); r.Add('PREDMET'); r.Add('O1'); r.Add('O2'); r.Add('YER'); r.Add('FL');
  v := TStringList.Create(); v.Clear();
  try
   for i := 2 to sg.RowCount-1 do
   begin
    // Бежим по предметам, перескакивая по 3 ячейки
     v.Clear();
     case tp of
      // по ученикам
      0: begin
          v.Add(IntToStr(integer(sg.Objects[0, 0])));
          v.Add(IntToStr(integer(sg.Objects[1, i])));
         end;
      // по предметам
      1: begin
          v.Add(IntToStr(integer(sg.Objects[1, i])));
          v.Add(IntToStr(integer(sg.Objects[0, 0])));
         end;
     end;

     // проверка
     t := sg.Cells[2, i];
     k := pos(',', t);
     if (k <> 0) then
     begin
      delete(t, k-1, 1);
      t := copy(t, 1, k-1) + '.' + copy(t, k+1, Length(t));
     end;
     sg.Cells[2, i] := t;
     t := sg.Cells[3, i];
     k := pos(',', t);
     if (k <> 0) then
     begin
      delete(t, k-1, 1);
      t := copy(t, 1, k-1) + '.' + copy(t, k+1, Length(t));
     end;
     sg.Cells[3, i] := t;
     t := sg.Cells[4, i];
     k := pos(',', t);
     if (k <> 0) then
     begin
      delete(t, k-1, 1);
      t := copy(t, 1, k-1) + '.' + copy(t, k+1, Length(t));
     end;
     sg.Cells[4, i] := t;

     v.Add(''''+sg.Cells[2, i]+'''');
     v.Add(''''+sg.Cells[3, i]+'''');
     v.Add(''''+sg.Cells[4, i]+'''');
     if ((pos('.',sg.Cells[2, i]) = 0)and(pos('.',sg.Cells[3, i]) = 0)and(pos('.',sg.Cells[4, i]) = 0)) then v.Add('0') else v.Add('1');

     if ((sg.Objects[2, i] = nil)or(integer(sg.Objects[2, i]) = 0)) then
     begin
      if (cInserts('TB_MARKS',r,v) = 1) then Exception.Create('');
     end else
     begin
      if (cUpdates('TB_MARKS',r,v,'where ID='+IntToStr(integer(sg.Objects[2, i]))) = 1) then Exception.Create('');
     end;
   end;

  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
  r.Free; v.Free;
 end;
end;

function TfData.ExpMarksExcelAll(sg: TAdvStringGrid): byte;
var ExcelApp, Workbook: Variant;
    i, j: integer;
    tmp: string;
begin
 try
  Result := 0;
  // Создание Excel
  ExcelApp := CreateOleObject('Excel.Application');
  // Отключаем реакцию Excel на события, чтобы ускорить вывод информации
  ExcelApp.Application.EnableEvents := FALSE;
  ExcelApp.Application.DisplayAlerts := FALSE;
  //  Создаем Книгу (Workbook)
  Workbook := ExcelApp.WorkBooks.Add;
  // Делаем Excel невидимым
  ExcelApp.Visible := TRUE;
  ExcelApp.WorkBooks[1].WorkSheets[1].Name := cSelectS('NUM','TB_CLASS','where ID='+IntToStr(integer(sg.Objects[0, 0]))+' ')+'-'+cSelectS('NAME','TB_CLASS','where ID='+IntToStr(integer(sg.Objects[0, 0]))+' ');
  // Формируем шапку
  ExcelApp.ActiveWorkBook.WorkSheets[1].Columns[1].ColumnWidth := 25;
  ExcelApp.Range['A1:A2'].Mergecells := True;
  tmp := sg.Cells[1, 1];
  delete(tmp,1,18);
  delete(tmp,pos('<',tmp),Length(tmp));
  ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, 1] := tmp;
  ExcelApp.Range['A1:A2'].HorizontalAlignment := 3;
  ExcelApp.Range['A1:A2'].VerticalAlignment := 2;

  i := 2;
  // Переписываем шапку
  while (i < sg.ColCount) do
  begin
   // названия предметов
   ExcelApp.Range[GetLabel(i)+'1'+':'+GetLabel(i+2)+'1'].Mergecells := True;
   tmp := sg.Cells[i, 1];
   delete(tmp,1,18);
   delete(tmp,pos('<',tmp),Length(tmp));
   ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, i] := tmp;
   ExcelApp.Range[GetLabel(i)+'1'+':'+GetLabel(i+3)+'1'].HorizontalAlignment := 3;
   // 1 сем.
   tmp := sg.Cells[i, 2];
   delete(tmp,1,18);
   delete(tmp,pos('<',tmp),Length(tmp));
   ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[2, i] := tmp;
   ExcelApp.Range[GetLabel(i)+'2'+':'+GetLabel(i)+'2'].HorizontalAlignment := 3;
   // 2 сем.
   tmp := sg.Cells[i+1, 2];
   delete(tmp,1,18);
   delete(tmp,pos('<',tmp),Length(tmp));
   ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[2, i+1] := tmp;
   ExcelApp.Range[GetLabel(i+1)+'2'+':'+GetLabel(i+1)+'2'].HorizontalAlignment := 3;
   // Год
   tmp := sg.Cells[i+2, 2];
   delete(tmp,1,18);
   delete(tmp,pos('<',tmp),Length(tmp));
   ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[2, i+2] := tmp;
   ExcelApp.Range[GetLabel(i+2)+'2'+':'+GetLabel(i+2)+'2'].HorizontalAlignment := 3;
   i := i+3;
  end;
  ExcelApp.Range['A1:'+GetLabel(i-1)+'2'].Borders.Weight := 2;

  // переписываем всё оставшиеся
  for i := 3 to sg.RowCount-1 do
   for j := 1 to sg.ColCount-1 do
    if (j = 1) then
     // Ф.И.О.
     ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[i, j]  := sg.Cells[j, i]
    else
     // Если в оценках пусто
     if (sg.Cells[j, i] = '') then ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[i, j] := 0 else
     // Если записано через запятую, элсе - точка
     if (pos('.',sg.Cells[j, i]) = 0) then
      ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[i, j]  := StrToFloat(sg.Cells[j, i])
     else
      ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[i, j]  := StrToFloat(copy(sg.Cells[j, i],1,pos('.',sg.Cells[j, i])-1)+','+copy(sg.Cells[j, i],pos('.',sg.Cells[j, i])+1,Length(sg.Cells[j, i])));

  ExcelApp.Range['A1:A'+IntToStr(i-1)].Borders.Weight := 2;
  // Делаем Excel видимым
  ExcelApp.Application.EnableEvents := TRUE;
  ExcelApp.Application.DisplayAlerts := TRUE;
  ExcelApp.Visible := TRUE;
 except
  Result := 1;
  ShowError(4);
 end;
end;

function TfData.GetLabel(Num: integer): string;
begin
 if (Num < 27) then Result := Labels[Num-1] else
 begin
 Result := Labels[(Num div 26)-1];
 if (((Num mod 26) = 0)and(Num > 26)) then Result := Labels[(Num div 26)-2]+'Z' else
 if (((Num mod 26) = 0)and(Num <= 26)) then Result := Labels[(Num div 26)-1]+'Z' else
 Result := Result + Labels[(Num mod 26)-1];
 end;
end;

function TfData.ExpMarksExcelP(tp: byte; sg: TAdvStringGrid): byte;
var ExcelApp, Workbook: Variant;
    i, j: integer;
    tmp: string;
begin
 try
  Result := 0;
  // Создание Excel
  ExcelApp := CreateOleObject('Excel.Application');
  // Отключаем реакцию Excel на события, чтобы ускорить вывод информации
  ExcelApp.Application.EnableEvents := FALSE;
  ExcelApp.Application.DisplayAlerts := FALSE;
  //  Создаем Книгу (Workbook)
  Workbook := ExcelApp.WorkBooks.Add;
  // Делаем Excel невидимым
  ExcelApp.Visible := FALSE;
  case tp of
   0: ExcelApp.WorkBooks[1].WorkSheets[1].Name := cSelectS('FNAME','TB_PEOP','where ID='+IntToStr(integer(sg.Objects[0, 0])))+' '+cSelectS('NAME','TB_PEOP','where ID='+IntToStr(integer(sg.Objects[0, 0])))+' '+cSelectS('SNAME','TB_PEOP','where ID='+IntToStr(integer(sg.Objects[0, 0])));
   1: ExcelApp.WorkBooks[1].WorkSheets[1].Name := cSelectS('NAME','TB_PREDMET','where ID='+IntToStr(integer(sg.Objects[0, 0])));
  end;
  // Формируем шапку
  ExcelApp.ActiveWorkBook.WorkSheets[1].Columns[1].ColumnWidth := 25;
  tmp := sg.Cells[1, 1];
  delete(tmp,1,18);
  delete(tmp,pos('<',tmp),Length(tmp));
  ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, 1] := tmp;
  ExcelApp.Range['A1:A1'].HorizontalAlignment := 3;

  // 1 сем.
   tmp := sg.Cells[2, 1];
   delete(tmp,1,18);
   delete(tmp,pos('<',tmp),Length(tmp));
   ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, 2] := (tmp);
   ExcelApp.Range['B1:B1'].HorizontalAlignment := 3;
   // 2 сем.
   tmp := sg.Cells[3, 1];
   delete(tmp,1,18);
   delete(tmp,pos('<',tmp),Length(tmp));
   ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, 3] := (tmp);
   ExcelApp.Range['C1:C1'].HorizontalAlignment := 3;
   // Год
   tmp := sg.Cells[4, 1];
   delete(tmp,1,18);
   delete(tmp,pos('<',tmp),Length(tmp));
   ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, 4] := (tmp);
   ExcelApp.Range['D1:D1'].HorizontalAlignment := 3;

   ExcelApp.Range['A1:D1'].Borders.Weight := 2;

  // переписываем всё оставшиеся
  for i := 2 to sg.RowCount-1 do
   for j := 1 to sg.ColCount-1 do
    if (j = 1) then
     // Ф.И.О.
     ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[i, j]  := sg.Cells[j, i]
    else
     // Если в оценках пусто
     if (sg.Cells[j, i] = '') then ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[i, j] := 0 else
     // Если записано через запятую, элсе - точка
     if (pos('.',sg.Cells[j, i]) = 0) then
      ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[i, j]  := StrToFloat(sg.Cells[j, i])
     else
      ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[i, j]  := StrToFloat(copy(sg.Cells[j, i],1,pos('.',sg.Cells[j, i])-1)+','+copy(sg.Cells[j, i],pos('.',sg.Cells[j, i])+1,Length(sg.Cells[j, i])));

  ExcelApp.Range['A1:A'+IntToStr(i-1)].Borders.Weight := 2;
  // Делаем Excel видимым
  ExcelApp.Application.EnableEvents := TRUE;
  ExcelApp.Application.DisplayAlerts := TRUE;
  ExcelApp.Visible := TRUE;
 except
  Result := 1;
  ShowError(4);
 end;
end;

function TfData.ExpMarksDiag(tp: byte; sg: TAdvStringGrid): byte;
var ExcelApp, Workbook, WorkSheet, Chart: Variant;
    i, j, m: integer;
    t1, t2: real;
begin
 Result := 0;
 try
  // Создание Excel
  ExcelApp := CreateOleObject('Excel.Application');
  // Отключаем реакцию Excel на события, чтобы ускорить вывод информации
  ExcelApp.Application.EnableEvents := FALSE;
  ExcelApp.Application.DisplayAlerts := FALSE;
  // Делаем Excel НЕвидимым
  ExcelApp.Visible := FALSE;
  //  Создаем Книгу (Workbook) с данными
  Workbook := ExcelApp.WorkBooks.Add;
  ExcelApp.WorkBooks[1].WorkSheets[1].Name := fData.standarts[20];
  WorkSheet := Workbook.WorkSheets[1];
  // Заполняем данные для диаграммы
  case tp of
   0: i := 2;
   1: i := 1;
   2: i := 1;
  end;
  j := 1;
   case tp of
    // Все на все
    0: while (i < sg.ColCount) do
       begin
        // 1 сем.
        ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, j] := GetNormalName(sg.Cells[i, 1]) + ' ' + GetNormalName(sg.Cells[i, 2]);
        ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[2, j] := StrToFloat(cSelectS('(sum(O1)/(select count(ID) from TB_PEOP where CLASS='+IntToStr(integer(sg.Objects[0,0]))+'))', 'TB_MARKS', 'where (PEOP in (select ID from TB_PEOP where CLASS='+IntToStr(integer(sg.Objects[0,0]))+'))and(PREDMET='+IntToStr(integer(sg.Objects[i+2,1]))+')'));
        // 2 сем.
        ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, j+1] := GetNormalName(sg.Cells[i, 1]) + ' ' + GetNormalName(sg.Cells[i+1, 2]);
        ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[2, j+1] := StrToFloat(cSelectS('(sum(O2)/(select count(ID) from TB_PEOP where CLASS='+IntToStr(integer(sg.Objects[0,0]))+'))', 'TB_MARKS', 'where (PEOP in (select ID from TB_PEOP where CLASS='+IntToStr(integer(sg.Objects[0,0]))+'))and(PREDMET='+IntToStr(integer(sg.Objects[i+2,1]))+')'));
        // Год
        ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, j+2] := GetNormalName(sg.Cells[i, 1]) + ' ' + GetNormalName(sg.Cells[i+2, 2]);
        ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[2, j+2] := StrToFloat(cSelectS('(sum(YER)/(select count(ID) from TB_PEOP where CLASS='+IntToStr(integer(sg.Objects[0,0]))+'))', 'TB_MARKS', 'where (PEOP in (select ID from TB_PEOP where CLASS='+IntToStr(integer(sg.Objects[0,0]))+'))and(PREDMET='+IntToStr(integer(sg.Objects[i+2,1]))+')'));
        j := j+3;
        i := i+3;
       end;
    // по ученикам
    1: for i := 2 to sg.RowCount -1 do
       begin
        // 1 сем.
        ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, j] := sg.Cells[1, i] + ' ' + GetNormalName(sg.Cells[2, 1]);
        // Если в оценках пусто
        if (sg.Cells[2, i] = '') then ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[2, j] := 0 else
        // Если записано через запятую, элсе - точка
        if (pos('.',sg.Cells[2, i]) = 0) then
         ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[2, j]  := StrToFloat(sg.Cells[2, i])
        else
         ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[2, j]  := StrToFloat(copy(sg.Cells[2, i],1,pos('.',sg.Cells[2, i])-1)+','+copy(sg.Cells[2, i],pos('.',sg.Cells[2, i])+1,Length(sg.Cells[2, i])));


        // 2 сем.
        ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, j+1] := sg.Cells[1, i] + ' ' + GetNormalName(sg.Cells[3, 1]);
        // Если в оценках пусто
        if (sg.Cells[3, i] = '') then ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[2, j+1] := 0 else
        // Если записано через запятую, элсе - точка
        if (pos('.',sg.Cells[3, i]) = 0) then
         ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[2, j+1]  := StrToFloat(sg.Cells[3, i])
        else
         ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[2, j+1]  := StrToFloat(copy(sg.Cells[3, i],1,pos('.',sg.Cells[3, i])-1)+','+copy(sg.Cells[3, i],pos('.',sg.Cells[3, i])+1,Length(sg.Cells[3, i])));


        // Год
        ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, j+2] := sg.Cells[1, i] + ' ' + GetNormalName(sg.Cells[4, 1]);
        // Если в оценках пусто
        if (sg.Cells[4, i] = '') then ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[2, j+2] := 0 else
        // Если записано через запятую, элсе - точка
        if (pos('.',sg.Cells[4, i]) = 0) then
         ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[2, j+2]  := StrToFloat(sg.Cells[4, i])
        else
         ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[2, j+2]  := StrToFloat(copy(sg.Cells[4, i],1,pos('.',sg.Cells[4, i])-1)+','+copy(sg.Cells[4, i],pos('.',sg.Cells[4, i])+1,Length(sg.Cells[4, i])));

        j := j+3;
       end;
    // по предметам
    2: begin
        // 1 сем.
        ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, 1] := cSelectS('NAME','TB_PREDMET','where ID='+IntToStr(integer(sg.Objects[0,0]))) + ' ' + GetNormalName(sg.Cells[2, 1]);
        ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[2, 1] := StrToFloat(cSelectS('(sum(O1)/(select count(ID) from TB_PEOP where CLASS='+IntToStr(integer(sg.Objects[1,1]))+'))', 'TB_MARKS', 'where (PEOP in (select ID from TB_PEOP where CLASS='+IntToStr(integer(sg.Objects[1,1]))+'))and(PREDMET='+IntToStr(integer(sg.Objects[0,0]))+')'));
        // 2 сем.
        ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, 2] := cSelectS('NAME','TB_PREDMET','where ID='+IntToStr(integer(sg.Objects[0,0]))) + ' ' + GetNormalName(sg.Cells[3, 1]);
        ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[2, 2] := StrToFloat(cSelectS('(sum(O2)/(select count(ID) from TB_PEOP where CLASS='+IntToStr(integer(sg.Objects[1,1]))+'))', 'TB_MARKS', 'where (PEOP in (select ID from TB_PEOP where CLASS='+IntToStr(integer(sg.Objects[1,1]))+'))and(PREDMET='+IntToStr(integer(sg.Objects[0,0]))+')'));
        // Год
        ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, 3] := cSelectS('NAME','TB_PREDMET','where ID='+IntToStr(integer(sg.Objects[0,0]))) + ' ' + GetNormalName(sg.Cells[4, 1]);
        ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[2, 3] := StrToFloat(cSelectS('(sum(YER)/(select count(ID) from TB_PEOP where CLASS='+IntToStr(integer(sg.Objects[1,1]))+'))', 'TB_MARKS', 'where (PEOP in (select ID from TB_PEOP where CLASS='+IntToStr(integer(sg.Objects[1,1]))+'))and(PREDMET='+IntToStr(integer(sg.Objects[0,0]))+')'));
        j := 4;
       end;

   end;
  //  Создаем Книгу (Workbook) с диаграммой
  Chart := WorkBook.Sheets.Add(,,1,xlChart);
  // Устанавливаем настройки диаграммы
  Chart.HasTitle :=1;
  case tp of
   0: Chart.ChartTitle.Text := cSelectS('NUM','TB_CLASS','where ID='+IntToStr(integer(sg.Objects[0,0]))) + '-' + cSelectS('NAME','TB_CLASS','where ID='+IntToStr(integer(sg.Objects[0,0])));
   1: Chart.ChartTitle.Text := cSelectS('FNAME','TB_PEOP','where ID='+IntToStr(integer(sg.Objects[0,0]))) + ' ' + cSelectS('NAME','TB_PEOP','where ID='+IntToStr(integer(sg.Objects[0,0]))) + ' ' + cSelectS('SNAME','TB_PEOP','where ID='+IntToStr(integer(sg.Objects[0,0])));
   2: Chart.ChartTitle.Text := cSelectS('NAME','TB_PREDMET','where ID='+IntToStr(integer(sg.Objects[0,0])))
  end;
  Chart.Axes(1).HasTitle := True;
  Chart.Axes(1).AxisTitle.Text := standarts[8];
  Chart.Axes(2).HasTitle := True;
  Chart.Axes(2).AxisTitle.Text := standarts[22];
  // Указываем диаграмме откуда брать данные
  Chart.SetSourceData(WorkSheet.Range['A1:'+GetLabel(j-1)+'2'], xlColumns);
  // Отключаем и задаем вручную  настройки осей
   Chart.Axes(2).MinimumScaleIsAuto := False;
   Chart.Axes(2).MaximumScaleIsAuto := False;
   Chart.Axes(2).MinimumScale:=0;

  // Задаем максимум оси. Если нет больше 12, то 12 элсе макс+1
  m := 12;
  if (tp = 0) then
  begin
   for i := 3 to sg.RowCount-1 do
   begin
    j := 5;
    while (j <= sg.ColCount) do
    begin
     // Если в оценках пусто
     if (sg.Cells[j-3, i] = '') then t1 := 0 else
     // Если записано через запятую, элсе - точка
     if (pos('.',sg.Cells[j-3, i]) = 0) then t1 := StrToFloat(sg.Cells[j-3, i])
     else t1 := StrToFloat(copy(sg.Cells[j-3, i],1,pos('.',sg.Cells[j-3, i])-1)+','+copy(sg.Cells[j-3, i],pos('.',sg.Cells[j-3, i])+1,Length(sg.Cells[j-3, i])));

     // Если в оценках пусто
     if (sg.Cells[j-2, i] = '') then t2 := 0 else
     // Если записано через запятую, элсе - точка
     if (pos('.',sg.Cells[j-2, i]) = 0) then t2 := StrToFloat(sg.Cells[j-2, i])
     else t2 := StrToFloat(copy(sg.Cells[j-2, i],1,pos('.',sg.Cells[j-2, i])-1)+','+copy(sg.Cells[j-2, i],pos('.',sg.Cells[j-2, i])+1,Length(sg.Cells[j-2, i])));

     m := Max(m, Round(t1));
     m := Max(m, Round(t2));

     // Если в оценках пусто
     if (sg.Cells[j-1, i] = '') then t1 := 0 else
     // Если записано через запятую, элсе - точка
     if (pos('.',sg.Cells[j-1, i]) = 0) then t1 := StrToFloat(sg.Cells[j-1, i])
     else t1 := StrToFloat(copy(sg.Cells[j-1, i],1,pos('.',sg.Cells[j-1, i])-1)+','+copy(sg.Cells[j-1, i],pos('.',sg.Cells[j-1, i])+1,Length(sg.Cells[j-1, i])));

     m := Max(m, Round(t1));
     j := j+3;
    end;

   end;
  end else
   for i := 2 to sg.RowCount-1 do
   begin
         // Если в оценках пусто
     if (sg.Cells[2, i] = '') then t1 := 0 else
     // Если записано через запятую, элсе - точка
     if (pos('.',sg.Cells[2, i]) = 0) then t1 := StrToFloat(sg.Cells[2, i])
     else t1 := StrToFloat(copy(sg.Cells[2, i],1,pos('.',sg.Cells[2, i])-1)+','+copy(sg.Cells[2, i],pos('.',sg.Cells[2, i])+1,Length(sg.Cells[2, i])));

     // Если в оценках пусто
     if (sg.Cells[3, i] = '') then t2 := 0 else
     // Если записано через запятую, элсе - точка
     if (pos('.',sg.Cells[3, i]) = 0) then t2 := StrToFloat(sg.Cells[3, i])
     else t2 := StrToFloat(copy(sg.Cells[3, i],1,pos('.',sg.Cells[3, i])-1)+','+copy(sg.Cells[3, i],pos('.',sg.Cells[3, i])+1,Length(sg.Cells[3, i])));

     m := Max(m, Round(t1));
     m := Max(m, Round(t2));

     // Если в оценках пусто
     if (sg.Cells[4, i] = '') then t1 := 0 else
     // Если записано через запятую, элсе - точка
     if (pos('.',sg.Cells[4, i]) = 0) then t1 := StrToFloat(sg.Cells[4, i])
     else t1 := StrToFloat(copy(sg.Cells[4, i],1,pos('.',sg.Cells[4, i])-1)+','+copy(sg.Cells[4, i],pos('.',sg.Cells[4, i])+1,Length(sg.Cells[4, i])));

     m := Max(m, Round(t1));
   end;

  if (m > 12) then m := m+1 else m := 12;
  Chart.Axes(2).MaximumScale:=m;

  Chart.Axes(2).MajorUnit := 1;
  // Вкключаем реакцию Excel на события
  ExcelApp.Application.EnableEvents := TRUE;
  ExcelApp.Application.DisplayAlerts := TRUE;
  // Делаем Excel видимым
  ExcelApp.Visible := TRUE;
 except
  Result := 1;
  ShowError(4);
 end;
end;

function TfData.GetNormalName(s: string): string;
var tmp: string;
begin
 tmp := s;
 delete(tmp,1,18);
 delete(tmp,pos('<',tmp),Length(tmp));
 Result := tmp;
end;

function TfData.ClassO(F, S: integer): byte;
var p: array of integer;
    i: integer;
begin
 try
  Result := 0;
  try
   i := 0;
   // Устанавливаем размер массива для идишников учеников
   SetLength(p,cCount('ID','TB_PEOP','where CLASS='+IntToStr(S)));
   SQL.SQL.Add('Select ID from TB_PEOP where CLASS='+IntToStr(S));
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   // Записываем в массив идишники
   while not(SQL.Eof) do
   begin
    p[i] := SQL.FieldByName('ID').AsInteger;
    Inc(i);
    SQL.Next();
   end;
   SQL.SQL.Clear();
   SQL.Close();
   Transaction.Active := FALSE;

   // Переводим в другой класс
   for i := 0 to Length(p)-1 do
    cUpdate('TB_PEOP', 'CLASS', IntToStr(F), 'Where ID='+IntToStr(p[i]));
   SetLength(p,0);

   ClassDel(S);
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.ClassDel(ID: integer): byte;
var p: array of integer;
    i: integer;
begin
 try
  Result := 0;
  try
   i := 0;
   // Устанавливаем размер массива для идишников учеников
   SetLength(p,cCount('ID','TB_PEOP','where CLASS='+IntToStr(ID)));
   SQL.SQL.Add('Select ID from TB_PEOP where CLASS='+IntToStr(ID));
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   // Записываем в массив идишники
   while not(SQL.Eof) do
   begin
    p[i] := SQL.FieldByName('ID').AsInteger;
    Inc(i);
    SQL.Next();
   end;
   SQL.SQL.Clear();
   SQL.Close();
   Transaction.Active := FALSE;

   // По очередно удаляем учеников
   for i := 0 to Length(p)-1 do PeopDel(p[i]);
   SetLength(p,0);

   if (cDelete('TB_CLASS','where ID='+IntToStr(ID)) = 1) then Exception.Create('');
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.PeopDel(ID: integer): byte;
begin
 if ((cDelete('TB_MARKS','where PEOP='+IntToStr(ID)) = 0)and(cDelete('TB_AMARKS','where PEOP='+IntToStr(ID)) = 0)and(cDelete('TB_MEDICINA','where ID='+IntToStr(ID)) = 0)and(cDelete('TB_PEOP','where ID='+IntToStr(ID)) = 0)) then Result := 0 else Result := 1;
end;

function TfData.MarksInArhiv(ID: integer): byte;
var i: integer;
    a: array of integer;
    p: array of Marks;
    r,v: TStringList;
begin
 try
  Result := 0;
  r := TStringList.Create(); r.Clear();
  v := TStringList.Create(); v.Clear();
  try
   i := 0;
   SetLength(a, 0);
   // Выбираем учеников класса
   SQL.SQL.Add('Select ID from TB_PEOP where CLASS='+IntToStr(ID));
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   // Записываем в массив идишники
   while not(SQL.Eof) do
   begin
    SetLength(a, Length(a)+1);
    a[Length(a)-1] := SQL.FieldByName('ID').AsInteger;
    Inc(i);
    SQL.Next();
   end;
   SQL.SQL.Clear();
   SQL.Close();
   Transaction.Active := FALSE;

   SetLength(p, 0);

   // Выбираем и записываем оценки учеников
   for i := 0 to Length(a)-1 do
   begin
    SQL.SQL.Add('select TB_PREDMET.NAME AS PREDMET, TB_CLASS.NUM AS CN, TB_TEACHER.FNAME AS TF, TB_TEACHER.NAME AS TN, TB_TEACHER.SNAME AS TS, TB_MARKS.FL, TB_MARKS.O1, TB_MARKS.O2, TB_MARKS.YER from TB_PREDMET, TB_TEACHER, TB_MARKS, TB_PLANCLASS, TB_UPLAN, TB_CLASS ');
    SQL.SQL.Add('where ((TB_MARKS.PEOP='+IntToStr(a[i])+')and(TB_CLASS.NUM<'+IntToStr(MaxClass)+')and(TB_CLASS.ID=(select CLASS from TB_PEOP where ID='+IntToStr(a[i])+'))and(TB_MARKS.PREDMET=TB_PREDMET.ID)and(TB_UPLAN.ID=TB_CLASS.UPLAN)and(TB_UPLAN.ID=TB_PLANCLASS.UPLAN)and(TB_TEACHER.ID=TB_PLANCLASS.TEACHER)and(TB_PLANCLASS.PREDMET=TB_PREDMET.ID))');
    Transaction.Active := TRUE;
    SQL.ExecQuery();
    // Записываем в массив данные
    while not(SQL.Eof) do
    begin
     SetLength(p, Length(p)+1);
     p[Length(p)-1].Peop := a[i];
     p[Length(p)-1].Clas := SQL.FieldByName('CN').AsInteger;
     p[Length(p)-1].Clas := SQL.FieldByName('CN').AsInteger;
     p[Length(p)-1].Predmet := SQL.FieldByName('PREDMET').AsString;
     p[Length(p)-1].Teacher := SQL.FieldByName('TF').AsString +' '+ SQL.FieldByName('TN').AsString[1] +'. '+ SQL.FieldByName('TS').AsString[1] +'.';
     p[Length(p)-1].O1 := SQL.FieldByName('O1').AsString;
     p[Length(p)-1].O2 := SQL.FieldByName('O2').AsString;
     p[Length(p)-1].YER := SQL.FieldByName('YER').AsString;
     p[Length(p)-1].FL := SQL.FieldByName('FL').AsInteger;
     SQL.Next();
    end;
    SQL.SQL.Clear();
    SQL.Close();
    Transaction.Active := FALSE;
   end;
   SQL.SQL.Clear();
   SQL.Close();
   Transaction.Active := FALSE;

    // Бежим по нашему массивчику и добавляем оценки в архив
   r.Add('PEOP'); r.Add('CLASS'); r.Add('PREDMET'); r.Add('TEACHER'); r.Add('O1'); r.Add('O2'); r.Add('YER'); r.Add('FL');
   for i := 0 to Length(p)-1 do
   begin
    v.Clear();
    v.Add(IntToStr(p[i].Peop)); v.Add(IntToStr(p[i].Clas));
    v.Add(''''+p[i].Predmet+''''); v.Add(''''+p[i].Teacher+'''');
    v.Add(''''+p[i].O1+''''); v.Add(''''+p[i].O2+''''); v.Add(''''+p[i].YER+''''); v.Add(IntToStr(p[i].FL));
    if (cInserts('TB_AMARKS',r,v) = 1) then Exception.Create('!!!');
   end;
 //  if (cDelete('TB_CLASS','where ID='+IntToStr(ID)) = 1) then Exception.Create('');
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
  r.Free(); v.Free();
 end;
end;

function TfData.FillTabsText(Tab: TAdvOfficeTabSet; What, Table, Where, Order: string): byte;
begin
 try
  Result := 0;
  Tab.AdvOfficeTabs.Clear;
  try
   Tab.AdvOfficeTabs.Add();
   Tab.AdvOfficeTabs[0].Tag     := -1;
   Tab.AdvOfficeTabs[0].Caption := standarts[0];
   SQL.SQL.Add('Select ID,'+What+' from '+Table+' '+Where+' '+Order);
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    Tab.AdvOfficeTabs.Add();
    Tab.AdvOfficeTabs[Tab.AdvOfficeTabs.Count-1].Tag     := SQL.FieldByName('ID').AsInteger;
    Tab.AdvOfficeTabs[Tab.AdvOfficeTabs.Count-1].Caption := SQL.FieldByName(What).AsString;
    SQL.Next();
   end;
   if (Tab.AdvOfficeTabs.Count > 1) then Tab.AdvOfficeTabs.Items[0].Free else Result := 1;
   Tab.ActiveTabIndex := 0;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.GoNextYear: byte;
var classes, peop: array of integer;
    i,z,j: integer;
    y,m,d: Word;
    tmp: string;
begin
 try
  Result := 0;
  SetLength(classes, 0);
  try
   // Выбираем все классы меньше максимального и загоняем АйДишники в массив
   SQL.SQL.Add('Select TB_CLASS.ID from TB_CLASS where NUM<'+IntToStr(MaxClass));
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    SetLength(classes, Length(classes)+1);
    classes[Length(classes)-1] := SQL.FieldByName('ID').AsInteger;
    SQL.Next();
   end;
   SQL.SQL.Clear();
   SQL.Close();
   Transaction.Active := FALSE;

   // Бежим по массиву классов и для каждого класса вызываем
   // функцию сохранения оценок и увеличиваем номер класса
   for i := 0 to Length(classes)-1 do
   begin
    // Переводим текущие оценки в архив
    if (MarksInArhiv(classes[i]) = 1) then Exception.Create('!!!');
    // Проверям не попадает ли номер клаасса на "прыгающий" и переводим класс
    z := StrToInt(cSelectS('NUM','TB_CLASS','where ID='+IntToStr(classes[i])));
    Inc(z);
    // Проверка на "прыгающий" номер класса
    if (z < Length(JumpClass)) then
      if (JumpClass[z] > 0) then z := JumpClass[z];
    // Проверка на окончание школы (11 класс)
    if (z >= MaxClass) then
    begin
     DecodeDate(Now(),y,m,d);
     tmp := cSelectS('NAME', 'TB_CLASS', 'where ID='+IntToStr(classes[i]));
     tmp := IntToStr(z-1)+'-'+tmp;
     if (cUpdate('TB_CLASS','NUM',IntToStr(y), 'where ID='+IntToStr(classes[i])) = 1) then Exception.Create('!!!');
     if (cUpdate('TB_CLASS','NAME',''''+tmp+'''', 'where ID='+IntToStr(classes[i])) = 1) then Exception.Create('!!!');
    end else
    if (cUpdate('TB_CLASS','NUM',IntToStr(z), 'where ID='+IntToStr(classes[i])) = 1) then Exception.Create('!!!');

    // Выбираем учеников класса
    SetLength(peop, 0);
    SQL.SQL.Add('Select TB_PEOP.ID from TB_PEOP where CLASS='+IntToStr(classes[i]));
    Transaction.Active := TRUE;
    SQL.ExecQuery();
    while not(SQL.Eof) do
    begin
     SetLength(peop, Length(peop)+1);
     peop[Length(peop)-1] := SQL.FieldByName('ID').AsInteger;
     SQL.Next();
    end;
    SQL.SQL.Clear();
    SQL.Close();
    Transaction.Active := FALSE;

    // Удаляем оценки учеников класса
    for j := 0 to Length(peop)-1 do
     if (cDelete('TB_MARKS','where PEOP ='+IntToStr(peop[j])) = 1) then Exception.Create('!!!');
   end;
  except

   Result := 1;
   ShowError(4);
  end;
 finally
  SetLength(classes, 0);
  SetLength(peop, 0);
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.FillClassEndSG(sg: TAdvStringGrid; Where, order: string): byte;
begin
 try
  Result := 0;
  sg.RowCount := 1;
  sg.Objects[0,0] := Pointer(-1);
  try
   SQL.SQL.Add('Select ID,NUM,NAME from TB_CLASS '+Trim(Where)+' '+Trim(order)+' ;');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    sg.Objects[0, sg.RowCount-1] := Pointer(SQL.FieldByName('ID').AsInteger);
    sg.Cells[0, sg.RowCount-1] := SQL.FieldByName('NUM').AsString+' '+SQL.FieldByName('NAME').AsString;
    sg.RowCount := sg.RowCount+1;
    SQL.Next();
   end;
   SQL.SQL.Clear();
   SQL.Close();
   Transaction.Active := FALSE;
   if ((sg.RowCount = 1)and(integer(sg.Objects[0,0]) = -1)) then Result := 1 else sg.RowCount := sg.RowCount-1;
   sg.Row := 0;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.FillMedAll(Clas: integer; sg: TAdvStringGrid; where: string): byte;
begin
 try
  Result := 0;
  sg.RowCount := 2;
  sg.ColCount := 4;
  sg.ColWidths[0] := 110;
  sg.Cells[0,0] := '<P align="center">'+standarts[1]+'</p>';
  sg.ColWidths[1] := 70;
  sg.Cells[1,0] := '<P align="center">'+standarts[25]+'</p>';
  sg.ColWidths[2] := 70;
  sg.Cells[2,0] := '<P align="center">'+standarts[26]+'</p>';
  sg.ColWidths[3] := sg.Width-270;
  sg.Cells[3,0] := '<P align="center">'+standarts[27]+'</p>';

  sg.Cells[0,1]:=''; sg.Cells[1,1]:=''; sg.Cells[2,1]:=''; sg.Cells[3,1]:='';

  sg.Objects[0,1] := Pointer(-1);
  try
   SQL.SQL.Add('Select TB_MEDICINA.ID,TB_MEDICINA.DB,TB_MEDICINA.DE,TB_MEDICINA.TXT, TB_PEOP.FNAME,TB_PEOP.NAME,TB_PEOP.SNAME from TB_MEDICINA,TB_PEOP where ((TB_PEOP.CLASS='+IntToStr(Clas)+')and(TB_MEDICINA.PEOP=TB_PEOP.ID))'+Trim(Where)+' order by TB_MEDICINA.DB desc ;');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    sg.Objects[0, sg.RowCount-1] := Pointer(SQL.FieldByName('ID').AsInteger);
    sg.Cells[0, sg.RowCount-1] := SQL.FieldByName('FNAME').AsString+' '+SQL.FieldByName('NAME').AsString[1]+'. '+SQL.FieldByName('SNAME').AsString[1]+'.';
    sg.Cells[1, sg.RowCount-1] := DateToStr(SQL.FieldByName('DB').AsDate);
    sg.Cells[2, sg.RowCount-1] := DateToStr(SQL.FieldByName('DE').AsDate);
    sg.Cells[3, sg.RowCount-1] := SQL.FieldByName('TXT').AsString;
    sg.RowCount := sg.RowCount+1;
    SQL.Next();
   end;
   SQL.SQL.Clear();
   SQL.Close();
   Transaction.Active := FALSE;
   if ((sg.RowCount = 2)and(integer(sg.Objects[0,1]) = -1)) then Result := 1 else sg.RowCount := sg.RowCount-1;
   sg.Row := 0;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.FillMed(ID: integer; sg: TAdvStringGrid; where: string): byte;
begin
 try
  Result := 0;
  sg.RowCount := 2;
  sg.ColCount := 3;
  sg.ColWidths[0] := 70;
  sg.Cells[0,0] := '<P align="center">'+standarts[25]+'</p>';
  sg.ColWidths[1] := 70;
  sg.Cells[1,0] := '<P align="center">'+standarts[26]+'</p>';
  sg.ColWidths[2] := sg.Width-160;
  sg.Cells[2,0] := '<P align="center">'+standarts[27]+'</p>';

  sg.Cells[0,1]:=''; sg.Cells[1,1]:=''; sg.Cells[2,1]:='';

  sg.Objects[0,1] := Pointer(-1);
  try
   SQL.SQL.Add('Select TB_MEDICINA.ID,TB_MEDICINA.DB,TB_MEDICINA.DE,TB_MEDICINA.TXT from TB_MEDICINA where (TB_MEDICINA.PEOP='+IntToStr(ID)+')'+Trim(Where)+' order by TB_MEDICINA.DB desc ;');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    sg.Objects[0, sg.RowCount-1] := Pointer(SQL.FieldByName('ID').AsInteger);
    sg.Cells[0, sg.RowCount-1] := DateToStr(SQL.FieldByName('DB').AsDate);
    sg.Cells[1, sg.RowCount-1] := DateToStr(SQL.FieldByName('DE').AsDate);
    sg.Cells[2, sg.RowCount-1] := SQL.FieldByName('TXT').AsString;
    sg.RowCount := sg.RowCount+1;
    SQL.Next();
   end;
   SQL.SQL.Clear();
   SQL.Close();
   Transaction.Active := FALSE;
   if ((sg.RowCount = 2)and(integer(sg.Objects[0,1]) = -1)) then Result := 1 else sg.RowCount := sg.RowCount-1;
   sg.Row := 0;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.FillClassSG(School: integer; sg: TAdvStringGrid; Where: string): byte;
begin
 try
  Result := 0;
  sg.RowCount := 2;
  sg.Cells[0,0] := '<P align="center">'+standarts[12]+'</P>';
  sg.Cells[0,1] := standarts[0];
  sg.Objects[0,1] := Pointer(-1);
  try
   SQL.SQL.Add('Select ID,NUM,NAME from TB_CLASS where (SCHOOL='+IntToStr(School)+') '+Trim(Where)+' order by NUM,NAME');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    sg.Objects[0, sg.RowCount-1] := Pointer(SQL.FieldByName('ID').AsInteger);
    sg.Cells[0, sg.RowCount-1] := SQL.FieldByName('NUM').AsString+'-'+SQL.FieldByName('NAME').AsString;
    sg.RowCount := sg.RowCount + 1;
    SQL.Next();
   end;
   if (sg.RowCount > 2) then sg.RowCount := sg.RowCount-1
    else if (integer(sg.Objects[0, 1]) = -1) then Result := 1;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
  sg.FixedRows := 1;
  sg.Row := 0;
 end;
end;

function TfData.ChangePlan(Clas, UPlan: integer): byte;
var peop: array of integer;
    i: integer;
begin
 try
  Result := 0;
  SetLength(peop, 0);
  // Спрашиваем уверенность
  if (MessageDlg(errors[13], mtConfirmation, [mbOk, mbCancel], 0) = mrOk) then
  try
   if (fData.cUpdate('TB_CLASS','UPLAN',IntToStr(UPlan), 'where ID='+IntToStr(Clas)) = 0) then
   begin
    // Выбираем учеников класса
    SQL.SQL.Add('Select TB_PEOP.ID from TB_PEOP where CLASS='+IntToStr(Clas));
    Transaction.Active := TRUE;
    SQL.ExecQuery();
    while not(SQL.Eof) do
    begin
     SetLength(peop, Length(peop)+1);
     peop[Length(peop)-1] := SQL.FieldByName('ID').AsInteger;
     SQL.Next();
    end;
    SQL.SQL.Clear();
    SQL.Close();
    Transaction.Active := FALSE;

    // Удаляем оценки учеников класса
    for i := 0 to Length(peop)-1 do
     if (cDelete('TB_MARKS','where PEOP ='+IntToStr(peop[i])) = 1) then Exception.Create('!!!');
   end;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SetLength(peop, 0);
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.FillAClassCb(cb: TAdvOfficeComboBox; ID: integer): byte;
begin
 try
  Result := 0;
  cb.Clear;
  cb.Items.Clear;
  cb.Items.AddObject(standarts[0], Pointer(-1));
  try
   SQL.SQL.Add('Select distinct CLASS from TB_AMARKS where PEOP='+IntToStr(ID)+' order by CLASS;');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    cb.Items.Add(SQL.FieldByName('CLASS').AsString);
    SQL.Next();
   end;
   if (cb.Items.Count > 1) then cb.Items.Delete(0) else Result := 1;
   cb.ItemIndex := 0;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.FillAMarks(ID, Clas: integer; sg: TAdvStringGrid): byte;
var i: integer;
begin
 try
  Result := 0;
  sg.SplitAllCells();
//  sg.
  sg.RowCount := 2;
  sg.ColCount := 5;
  sg.ColWidths[1] := 150;
  try
   // Формируем шапку
   { ----------------|-----------------------------|-------
                     |  1семестр | 2 семестр | Год |
     ---------------------------------------------------
     Ученик|Предмет  |    10     |    11     |  11 |
     ---------------------------------------------------
   }
   sg.Objects[0,0] := Pointer(ID);
   sg.Cells[1,1] := '<P align="center">'+standarts[18]+'</p>';
   // 1 сем., 2 сем., Год
   sg.Cells[2, 1] := '<P align="center">'+standarts[15]+'</p>';
   sg.Cells[3, 1] := '<P align="center">'+standarts[16]+'</p>';
   sg.Cells[4, 1] := '<P align="center">'+standarts[17]+'</p>';
//   sg.FixedRows := 2;
   sg.FixedCols := 1;
   sg.RowHeights[0] := -1;
   sg.ColWidths[0] := -1;

   // Выбираем предметы
   SQL.SQL.Add('Select PREDMET from TB_AMARKS where ((PEOP='+IntToStr(ID)+')and(CLASS='+IntToStr(Clas)+')) order by PREDMET');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    sg.RowCount := sg.RowCount+1;
    sg.Cells[1, sg.RowCount-1] := SQL.FieldByName('PREDMET').AsString;
    SQL.Next();
   end;
   SQL.SQL.Clear();
   SQL.Close();
   Transaction.Active := FALSE;

   // Выбираем оценки
   // Бежим по ученикам вниз
   for i := 2 to sg.RowCount-1 do
   begin
    SQL.SQL.Add('Select O1, O2, YER from TB_AMARKS where ((PEOP='+IntToStr(ID)+')and(CLASS='+IntToStr(Clas)+')and(PREDMET='''+sg.Cells[1, i]+'''))');
    Transaction.Active := TRUE;
    SQL.ExecQuery();
    if (SQL.FieldByName('O1').AsString = '') then sg.Cells[2, i] := '0' else
    sg.Cells[2, i]   := SQL.FieldByName('O1').AsString;
    if (SQL.FieldByName('O2').AsString = '') then sg.Cells[3, i] := '0' else
    sg.Cells[3, i]   := SQL.FieldByName('O2').AsString;
    if (SQL.FieldByName('YER').AsString = '') then sg.Cells[4, i] := '0' else
    sg.Cells[4, i]   := SQL.FieldByName('YER').AsString;
    SQL.SQL.Clear();
    SQL.Close();
    Transaction.Active := FALSE;
   end;

   if (sg.RowCount > 2) then sg.FixedRows := 2 else sg.FixedRows := 1;
   sg.FixedCols := 2;
   if (sg.RowCount = 2) then Result := 1;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear();
  SQL.Close();
  Transaction.Active := FALSE;
 end;
end;

function TfData.SrClassMark(ID: integer): real;
begin
 try
  Result := 0;
  try
   SQL.SQL.Add('select ((sum(TB_MARKS.YER)/count(TB_PEOP.ID))/count(distinct TB_MARKS.PREDMET)) as RS from TB_PEOP, TB_MARKS where ((TB_PEOP.CLASS='+IntToStr(ID)+')and(TB_PEOP.ID=TB_MARKS.PEOP)and(TB_MARKS.FL=0)) ;');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   Result := SQL.FieldByName('RS').AsFloat;
  except
   Result := -1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear;
  SQL.Close;
  Transaction.Active := FALSE;
 end;
end;

function TfData.SrSchoolMark(ID: integer): real;
begin
 try
  Result := 0;
  try
   SQL.SQL.Add('select ((sum(TB_MARKS.YER)/count(TB_PEOP.ID))/count(distinct TB_MARKS.PREDMET)) as RS from TB_PEOP, TB_MARKS, TB_CLASS where ((TB_CLASS.SCHOOL='+IntToStr(ID)+')and(TB_PEOP.CLASS=TB_CLASS.ID)and(TB_PEOP.ID=TB_MARKS.PEOP)and(TB_MARKS.FL=0)) ;');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   Result := SQL.FieldByName('RS').AsFloat;
  except
   Result := -1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear;
  SQL.Close;
  Transaction.Active := FALSE;
 end;
end;

function TfData.FillSGSchoolStat(sg: TAdvStringGrid; ID: integer): byte;
var i: integer;
    s: string;
begin
 try
  Result := 0;
  sg.ColCount := 3;
  sg.ColWidths[0] := 60;
  sg.ColWidths[1] := ((sg.Width-60) div 2)-2;
  sg.ColWidths[2] := ((sg.Width-60) div 2)-2;
  sg.Cells[0, 0] := '<P align="center">'+standarts[12]+'</P>';
  sg.Cells[1, 0] := '<P align="center">'+standarts[28]+'</P>';
  sg.Cells[2, 0] := '<P align="center">'+standarts[29]+'</P>';
  sg.RowCount := 1;
  try
   // выставляем номера паралелей
   SQL.SQL.Add('select distinct TB_CLASS.NUM AS RS from TB_CLASS where ((SCHOOL='+IntToStr(ID)+')and(NUM between 1 and '+IntToStr(MaxClass-1)+')) order by NUM;');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    sg.RowCount := sg.RowCount+1;
    sg.Cells[0, sg.RowCount-1] := IntToStr(SQL.FieldByName('RS').AsInteger);
    SQL.Next();
   end;
   SQL.SQL.Clear;
   SQL.Close;
   Transaction.Active := FALSE;

   // если нет классов - выходим
   if (sg.RowCount = 1) then exit;


   for i := 1 to sg.RowCount-1 do
   begin
    // заполняем количество учеников в паралелях
    SQL.SQL.Add('select count(TB_PEOP.ID) as RS from TB_PEOP, TB_CLASS where ((TB_PEOP.CLASS=TB_CLASS.ID)and(TB_CLASS.SCHOOL='+IntToStr(ID)+')and(TB_CLASS.NUM='+sg.Cells[0, i]+')) ;');
    Transaction.Active := TRUE;
    SQL.ExecQuery();
    sg.Cells[1, i] := IntToStr(SQL.FieldByName('RS').AsInteger);
    SQL.SQL.Clear;
    SQL.Close;
    Transaction.Active := FALSE;

    // заполняем среднюю оценку
    SQL.SQL.Add('select ((sum(TB_MARKS.YER)/count(TB_PEOP.ID))/count(distinct TB_MARKS.PREDMET)) as RS from TB_PEOP, TB_MARKS, TB_CLASS where ((TB_CLASS.SCHOOL='+IntToStr(ID)+')and(TB_CLASS.NUM='+sg.Cells[0, i]+')and(TB_PEOP.CLASS=TB_CLASS.ID)and(TB_PEOP.ID=TB_MARKS.PEOP)and(TB_MARKS.FL=0)) ;');
    Transaction.Active := TRUE;
    SQL.ExecQuery();
    s := FloatToStr(SQL.FieldByName('RS').AsFloat);
    if (Length(s) > 5) then s := copy(s, 1, 5);
    sg.Cells[2, i] := s;
    SQL.SQL.Clear;
    SQL.Close;
    Transaction.Active := FALSE;
   end;
   sg.FixedRows := 1;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear;
  SQL.Close;
  Transaction.Active := FALSE;
 end;
end;

function TfData.FillImg(img: TImage; ID: integer): byte;
var Path: string;
begin
 try
  Result := 0;
  try
   if (StrToInt(cSelectS('IMG','TB_SCHOOL','where ID='+IntToStr(ID))) = 0) then img.Picture := nil else
   begin
    Path := ExtractFileDir(Database.DatabaseName);
    if (Length(Path) = 0) then Path := ExtractFileDir(Application.ExeName) else
    if ((Path[1] <> '\')and(Path[2] <> ':')) then Path := ExtractFileDir(Application.ExeName);
    Path := Path + '\imgs\';
    img.Picture.LoadFromFile(Path+IntToStr(ID)+'.jpg');
   end;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear;
  SQL.Close;
  Transaction.Active := FALSE;
 end;
end;

function TfData.cUpdateImg(ID: integer; Table: string; img: TImage): byte;
begin
 try
  Result := 0;
  try
   if (img.Picture = nil) then
   begin
    if FileExists(ExtractFileDir(Application.ExeName)+'\imgs\'+IntToStr(ID)+'.jpg') then
     DeleteFile(ExtractFileDir(Application.ExeName)+'\imgs\'+IntToStr(ID)+'.jpg');
   end else img.Picture.SaveToFile(ExtractFileDir(Application.ExeName)+'\imgs\'+IntToStr(ID)+'.jpg');
   if (img.Picture = nil) then  SQL.SQL.Add('update '+Table+' set IMG = 0 where ID='+IntToStr(ID)+' ;') else
    SQL.SQL.Add('update '+Table+' set IMG = 1 where ID='+IntToStr(ID)+' ;');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   Transaction.Commit();
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear;
  SQL.Close;
  Transaction.Active := FALSE;
 end;
end;

function TfData.DeleteSchool(ID: integer): byte;
var classes: array of integer;
    i: integer;
begin
 try
  try
   Result := 0;
   // Выбираем все связанные со школой классы и убиваем их
   SetLength(classes, 0);
   SQL.SQL.Add('select ID from TB_CLASS where SCHOOL='+IntToStr(ID));
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   while not(SQL.Eof) do
   begin
    SetLength(classes, Length(classes)+1);
    classes[Length(classes)-1] := SQL.FieldByName('ID').AsInteger;
    SQL.Next;
   end;
   SQL.SQL.Clear;
   SQL.Close;
   Transaction.Active := FALSE;
   for i := 0 to Length(classes)-1 do
    ClassDel(classes[i]);
   // удаляем учителей
   cDelete('TB_TEACHER','where SCHOOL='+IntToStr(ID));
   // удаляем учебные планы
   cDelete('TB_PLANCLASS','where UPLAN in (select ID from TB_UPLAN where SCHOOL='+IntToStr(ID)+')');
   cDelete('TB_UPLAN','where SCHOOL='+IntToStr(ID));
   // удаляем саму школу и герб (если есть)
   if (StrToInt(cSelectS('IMG','TB_SCHOOL','where ID='+IntToStr(ID))) = 1) then
   if FileExists(ExtractFileDir(Application.ExeName)+'\imgs\'+IntToStr(ID)+'.jpg') then
     DeleteFile(ExtractFileDir(Application.ExeName)+'\imgs\'+IntToStr(ID)+'.jpg');
   cDelete('TB_SCHOOL','where ID='+IntToStr(ID));
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear;
  SQL.Close;
  Transaction.Active := FALSE;
  SetLength(classes, 0);
 end;
end;

function TfData.Find(sg: TAdvStringGrid; What: string): byte;
var f,n,s,t: string;
    i: integer;
begin
 try
  Result := 0;
  t := What;
  f := '';
  n := '';
  s := '';
  i := pos(' ',t);
  if (i <> 0) then
  begin
   f := copy(t,0,i-1);
   delete(t,0,i+1);
   i := pos(' ',t);

   if (Length(t) > 0) then
   begin
    i := pos(' ',t);
    if (i = 0) then n := t else
    begin
     n := copy(t,0,i-1);
     s := copy(t,i+1,Length(t));
    end;
   end;
  end else f := t;
  try
   sg.RowCount := 1;
   sg.Cells[0,0] := '<P align="center">'+standarts[1]+'</P>';
   sg.Cells[1,0] := '<P align="center">'+standarts[30]+'</P>';
   sg.Cells[2,0] := '<P align="center">'+standarts[12]+'</P>';
   SQL.SQL.Add('select TB_PEOP.ID as PID, TB_PEOP.FNAME as PF, TB_PEOP.NAME as PN, TB_PEOP.SNAME as PS, TB_SCHOOL.ID as SID, TB_SCHOOL.NAME as SN, TB_CLASS.ID as CID, TB_CLASS.NUM as CN, TB_CLASS.NAME as CM from TB_PEOP, TB_SCHOOL, TB_CLASS ');
   SQL.SQL.Add(' where ((TB_PEOP.CLASS=TB_CLASS.ID)and(TB_CLASS.SCHOOL=TB_SCHOOL.ID)and((TB_PEOP.FNAME LIKE ''%'+Trim(f)+'%'')and(TB_PEOP.NAME LIKE ''%'+Trim(n)+'%'')and(TB_PEOP.SNAME LIKE ''%'+Trim(s)+'%'')))');
   Transaction.Active := TRUE;
   SQL.ExecQuery();

   while not(SQL.Eof) do
   begin
    sg.RowCount := sg.RowCount+1;
    sg.Objects[0, sg.RowCount-1] := pointer(SQL.FieldByName('PID').AsInteger);
    sg.Cells[0, sg.RowCount-1] := SQL.FieldByName('PF').AsString +' '+ SQL.FieldByName('PN').AsString +' '+ SQL.FieldByName('PS').AsString;
    sg.Objects[1, sg.RowCount-1] := pointer(SQL.FieldByName('SID').AsInteger);
    sg.Cells[1, sg.RowCount-1] := SQL.FieldByName('SN').AsString;
    sg.Objects[2, sg.RowCount-1] := pointer(SQL.FieldByName('CID').AsInteger);
    sg.Cells[2, sg.RowCount-1] := SQL.FieldByName('CN').AsString +' '+ SQL.FieldByName('CM').AsString;

    SQL.Next();
   end;

   if (sg.RowCount > 1) then sg.FixedRows := 1;
  except
   Result := 1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear;
  SQL.Close;
  Transaction.Active := FALSE;
 end;
end;

function TfData.ExpMedExcel(sg: TAdvStringGrid): byte;
var ExcelApp, Workbook: Variant;
    i, j: integer;
    tmp: string;
begin
 try
  Result := 0;
  // Создание Excel
  ExcelApp := CreateOleObject('Excel.Application');
  // Отключаем реакцию Excel на события, чтобы ускорить вывод информации
  ExcelApp.Application.EnableEvents := FALSE;
  ExcelApp.Application.DisplayAlerts := FALSE;
  //  Создаем Книгу (Workbook)
  Workbook := ExcelApp.WorkBooks.Add;
  // Делаем Excel невидимым
  ExcelApp.Visible := FALSE;

  // Формируем шапку
  ExcelApp.ActiveWorkBook.WorkSheets[1].Columns[1].ColumnWidth := 25;
  tmp := sg.Cells[0, 0];
  delete(tmp,1,18);
  delete(tmp,pos('<',tmp),Length(tmp));
  ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, 1] := tmp;
  ExcelApp.Range['A1:A1'].HorizontalAlignment := 3;

   // нач. дата
   ExcelApp.ActiveWorkBook.WorkSheets[1].Columns[2].ColumnWidth := 15;
   tmp := sg.Cells[1, 0];
   delete(tmp,1,18);
   delete(tmp,pos('<',tmp),Length(tmp));
   ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, 2] := (tmp);
   ExcelApp.Range['B1:B1'].HorizontalAlignment := 3;
   // кон. дата
   ExcelApp.ActiveWorkBook.WorkSheets[1].Columns[3].ColumnWidth := 15;
   tmp := sg.Cells[2, 0];
   delete(tmp,1,18);
   delete(tmp,pos('<',tmp),Length(tmp));
   ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, 3] := (tmp);
   ExcelApp.Range['C1:C1'].HorizontalAlignment := 3;
   // Причина
   ExcelApp.ActiveWorkBook.WorkSheets[1].Columns[4].ColumnWidth := 100;
   tmp := sg.Cells[3, 0];
   delete(tmp,1,18);
   delete(tmp,pos('<',tmp),Length(tmp));
   ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[1, 4] := (tmp);
   ExcelApp.Range['D1:D1'].HorizontalAlignment := 3;

   ExcelApp.Range['A1:D1'].Borders.Weight := 2;

  // переписываем всё оставшиеся
  for i := 1 to sg.RowCount-1 do
   for j := 0 to sg.ColCount-1 do
     ExcelApp.ActiveWorkBook.WorkSheets[1].Cells[i+1, j+1]  := sg.Cells[j, i];

  // Делаем Excel видимым
  ExcelApp.Application.EnableEvents := TRUE;
  ExcelApp.Application.DisplayAlerts := TRUE;
  ExcelApp.Visible := TRUE;
 except
  Result := 1;
  ShowError(4);
 end;
end;

procedure TfData.SQLMonitorSQL(EventText: String; EventTime: TDateTime);
var s: string;
    fh: integer;
begin
 if not(FileExists(ExtractFileDir(Application.ExeName)+'\log.$$$')) then
 begin
  s := DateTimeToStr(EventTime)+'   '+EventText;
  fh := FileCreate(ExtractFileDir(Application.ExeName)+'\log.$$$');
  FileWrite(fh, s[1], Length(s));
  FileClose(fh);
 end else
 begin
  s := DateTimeToStr(EventTime)+'   '+EventText;
  // Проверяем размер файла и коль что очищаем его
  if (GetFileSize(ExtractFileDir(Application.ExeName)+'\log.$$$') > LengthLog) then DeleteFile(ExtractFileDir(Application.ExeName)+'\log.$$$');
  fh := FileOpen(ExtractFileDir(Application.ExeName)+'\log.$$$', fmOpenWrite);
  FileSeek(fh,0,2);
  FileWrite(fh, s[1], Length(s));
  FileClose(fh);
 end;
end;

function TfData.GetFileSize(Path: string): int64;
Var
  F: Integer;
Begin
  F:=FileOpen(Path,0);  { режим ReadOnly }
  Result := FileSeek(F,0,2);
  FileClose(F);
end;

function TfData.ChangeLanguage(Path: string): byte;
var f,e: TIniFile;
    s: string;
    i: integer;
    ers: TStringList;
begin
 try
  s := ExtractFileName(Path);
  delete(s, pos('.',s), Length(s));
  f := TIniFile.Create(ExtractFileDir(Application.ExeName)+'\config.ini');
  f.WriteString('Language', 'current', s);
  f.Free;
  language := s;

  // читаем и сохраняем сообщения об ошибках
  e := TIniFile.Create(ExtractFileDir(Application.ExeName)+'\language\'+language+'.ini');
  SetLength(errors,0);
  ers := TStringList.Create;
  e.ReadSection('ERRORS',ers);
  SetLength(errors,ers.Count);
  for i := 0 to ers.Count-1 do
   errors[i] := e.ReadString('ERRORS',IntToStr(i),'ERROR');
  // читаем и сохраняем стандартные надписи
  ers.Clear;
  e.ReadSection('STANDART',ers);
  SetLength(standarts,ers.Count);
  for i := 0 to ers.Count-1 do
   standarts[i] := e.ReadString('STANDART',IntToStr(i),'ERROR');
  ers.Free;
  e.Free;

  // изменяем текст на формах
  for i := 0 to Application.ComponentCount-1 do
   if ((Application.Components[i].ClassName[1] = 'T')and(Application.Components[i].ClassName[2] = 'f')and(Application.Components[i].ClassName <> 'TfData')) then
    SetLanguage(Application.Components[i] as TForm);
  SetLanguage(fMain);
  Result := 0;
 except
  Result := 1;
  ShowError(3);
 end;
end;

function TfData.ExpWord(t: byte; ID: integer): byte;
var FileName, cur: OleVariant;
    Table, s: string;
begin
 try
  case t of
   0 : begin
        s := 'Peop'+'_'+language;
        Table := 'TB_PEOP';
       end;
   1 : begin
        s := 'Teacher'+'_'+language;
        Table := 'TB_TEACHER';
       end;
  end;
  FileName:=ExtractFileDir(Application.ExeName)+'\Templates\'+s+'.dot';
 try  // Word не запущен, запустить
  WordA.Disconnect;
  WordA.Connect;
  WordA.Visible := TRUE;
//  WordA.Visible := FALSE;
 except
  WordA.Disconnect;
  ShowError(21);
  Exit;
 end;
 with WordA do
 begin
  WordA.Documents.Open(FileName,EmptyParam,EmptyParam,EmptyParam,
                          EmptyParam,EmptyParam,EmptyParam,
	                  EmptyParam,EmptyParam,EmptyParam,
	                  EmptyParam,EmptyParam,EmptyParam,EmptyParam,EmptyParam);
//  WordA.SelectFirst;
  Selection.NextField;
  while (Selection.Text <> 'q')or(Selection.Text <> 'Q') do
  begin
   case Selection.Text[1] of
    { ТЕКСТОВОЕ ПОЛЕ }
    't','T': Selection.Text := cSelectS(Copy(Selection.Text,2,Length(Selection.Text)), Table, 'where ID='+IntToStr(ID));
    { ДАТА }
    'd','D': begin
              s := cSelectS(Copy(Selection.Text,2,Length(Selection.Text)), Table, 'where ID='+IntToStr(ID));
              if (Trim(s) <> '') then s := DateToStr(StrToDate(s)) else s := '-';
              Selection.Text := s;
             end;
    { СПЕЦИАЛЬНЫЕ ПОЛЯ }
    's', 'S': case Selection.Text[2] of
               's', 'S': Selection.Text := cSelectS('TB_SCHOOL.NAME', 'TB_SCHOOL, TB_CLASS, TB_PEOP', 'where ((TB_SCHOOL.ID=TB_CLASS.SCHOOL)and(TB_CLASS.ID=TB_PEOP.CLASS)and(TB_PEOP.ID='+IntToStr(ID)+'))');
               'c', 'C': begin
                          s := cSelectS('TB_CLASS.NUM', 'TB_CLASS, TB_PEOP', 'where ((TB_CLASS.ID=TB_PEOP.CLASS)and(TB_PEOP.ID='+IntToStr(ID)+'))');
                          s := s+'-'+cSelectS('TB_CLASS.NAME', 'TB_CLASS, TB_PEOP', 'where ((TB_CLASS.ID=TB_PEOP.CLASS)and(TB_PEOP.ID='+IntToStr(ID)+'))');
                          Selection.Text := s;
                         end;
              end;
    { ВЫХОД }
    'q','Q': break;


   end;
   Selection.NextField;
  end;
end;
  WordA.Selection.Text := '';
  WordA.Selection.Delete(EmptyParam,EmptyParam);
  Result := 0;
 except
  ShowError(20);
  Result := 1;
 end;
end;

function TfData.CheckJump(cl: integer): byte;
var i: integer;
begin
 Result := 0;
 for i := 0 to Length(JumpClass)-1 do
  if (JumpClass[i] = cl) then
  begin
   Result := 1;
   break;
  end;
end;

function TfData.ifCount(What, Table, Where: string): integer;
begin
 try
  Result := 0;
  try
   SQL.SQL.Add('Select Count('+Trim(What)+') AS RS from '+Trim(Table)+' '+Trim(Where)+' ;');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   if (SQL.FieldByName('RS').AsInteger > 0) then Result := 1 else Result := 0;
  except
   Result := -1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear;
  SQL.Close;
  Transaction.Active := FALSE;
 end;
end;

function TfData.cMax(What, Table, Where: string): integer;
begin
  try
  Result := 0;
  try
   SQL.SQL.Add('Select MAX('+Trim(What)+') AS RS from '+Trim(Table)+' '+Trim(Where)+' ;');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   if (SQL.FieldByName('RS').AsInteger > 0) then Result := 1 else Result := 0;
  except
   Result := -1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear;
  SQL.Close;
  Transaction.Active := FALSE;
 end;
end;

function TfData.SrClassAMark(ID: integer): real;
begin
 try
  Result := 0;
  try
   SQL.SQL.Add('select ((sum(TB_AMARKS.YER)/count(TB_PEOP.ID))/count(distinct TB_AMARKS.PREDMET)) as RS from TB_PEOP, TB_AMARKS where ((TB_PEOP.CLASS='+IntToStr(ID)+')and(TB_PEOP.ID=TB_AMARKS.PEOP)and(TB_AMARKS.FL=0)) ;');
   Transaction.Active := TRUE;
   SQL.ExecQuery();
   Result := SQL.FieldByName('RS').AsFloat;
  except
   Result := -1;
   ShowError(4);
  end;
 finally
  SQL.SQL.Clear;
  SQL.Close;
  Transaction.Active := FALSE;
 end;
end;

end.
