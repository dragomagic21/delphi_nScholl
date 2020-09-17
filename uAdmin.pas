unit uAdmin;

interface

uses
  IniFiles, Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, AdvOfficePager, StdCtrls, AdvEdit, AdvGlassButton, Grids,
  BaseGrid, AdvGrid, AdvGroupBox, AdvOfficeButtons, ExtCtrls, AdvFontCombo;

type
  TfAdmin = class(TForm)
    Pager: TAdvOfficePager;
    pUsers: TAdvOfficePage;
    pDB: TAdvOfficePage;
    sgUsers: TAdvStringGrid;
    lLogin: TLabel;
    lPass: TLabel;
    bUserAdd: TAdvGlassButton;
    bUserDel: TAdvGlassButton;
    bUserEdit: TAdvGlassButton;
    eLogin: TAdvEdit;
    ePass: TAdvEdit;
    rgUsers: TAdvOfficeRadioGroup;
    eDB: TAdvEdit;
    lDB: TLabel;
    bBDSelect: TAdvGlassButton;
    bBDOk: TAdvGlassButton;
    OD: TOpenDialog;
    lDBWarning: TLabel;
    lDBLogin: TLabel;
    eDBLogin: TAdvEdit;
    lDBPass: TLabel;
    eDBPass: TAdvEdit;
    Bevel1: TBevel;
    pJump: TAdvOfficePage;
    sgJump: TAdvStringGrid;
    cbClassIn: TAdvOfficeComboBox;
    LClassIn: TLabel;
    cbClassOut: TAdvOfficeComboBox;
    LClassOut: TLabel;
    bClassJAdd: TAdvGlassButton;
    bClassJDel: TAdvGlassButton;
    procedure FormCreate(Sender: TObject);
    procedure PagerChange(Sender: TObject);
    procedure sgUsersClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bUserEditClick(Sender: TObject);
    procedure bUserAddClick(Sender: TObject);
    procedure bUserDelClick(Sender: TObject);
    procedure bBDSelectClick(Sender: TObject);
    procedure bBDOkClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure cbClassInChange(Sender: TObject);
    procedure bClassJAddClick(Sender: TObject);
    procedure bClassJDelClick(Sender: TObject);
  private
    m: byte;
    { Private declarations }
    procedure CreateParams(var Params: TCreateParams); override;
  public
    { Public declarations }
  end;

var
  fAdmin: TfAdmin;

implementation

uses uData;

{$R *.dfm}

procedure TfAdmin.FormCreate(Sender: TObject);
begin
 fData.SetLanguage(fAdmin);
 m := 0;
end;

procedure TfAdmin.PagerChange(Sender: TObject);
var f: TIniFile;
    i: integer;
begin
 case Pager.ActivePageIndex of
  // Пользователи
  0: begin
      case fData.cFillSg(sgUsers,31,'LOGIN','TB_USERS') of
       0: sgUsersClick(self);
       1: sgUsersClick(self);
      end;
     end;
  // Прыжки класса
  1: begin
      sgJump.Cells[0, 0] := '<P align="center">'+fData.standarts[33]+'</P>';
      sgJump.Cells[1, 0] := '<P align="center">'+fData.standarts[34]+'</P>';
      sgJump.RowCount := 2;
      cbClassIn.Items.Clear;
      for i := 1 to fData.MaxClass-1 do
       if ((fData.JumpClass[i] = 0)and(fData.CheckJump(i) = 0)) then
        // заполняем перечень классов под прыжок
        cbClassIn.Items.Add(IntToStr(i)) else
       if (fData.JumpClass[i] <> 0) then
       begin
        // заполняем список прыгающих классов
        sgJump.Cells[0, sgJump.RowCount-1] := IntToStr(i);
        sgJump.Cells[1, sgJump.RowCount-1] := IntToStr(fData.JumpClass[i]);
        sgJump.RowCount := sgJump.RowCount + 1;
       end;
      cbClassIn.ItemIndex := 0;
      cbClassInChange(self);
      sgJump.FixedRows := 1;
      sgJump.RowCount := sgJump.RowCount - 1;
      if (sgJump.RowCount = 1) then bClassJDel.Enabled := FALSE else bClassJDel.Enabled := TRUE;
      if ((cbClassIn.Items.Count = 0)or(cbClassOut.Items.Count = 0)) then bClassJAdd.Enabled := FALSE else bClassJAdd.Enabled := TRUE;
     end;
  // Настройки БД
  2: begin
      f := TIniFile.Create(ExtractFileDir(Application.ExeName)+'\config.ini');
      eDB.Text := f.ReadString('DataBase', 'Path', fData.errors[0]);
      eDBLogin.Text := f.ReadString('DataBase', 'user_name', fData.errors[0]);
      eDBPass.Text  := f.ReadString('DataBase', 'password', fData.errors[0]);
      f.Free;
     end;
 end;
end;

procedure TfAdmin.sgUsersClick(Sender: TObject);
begin
 if (m = 0) then exit;
 eLogin.Text := fData.cSelectS('LOGIN', 'TB_USERS', 'where ID='+IntToStr(integer(sgUsers.Objects[0, sgUsers.Row])));
 ePass.Text  := fData.cSelectS('PASS', 'TB_USERS', 'where ID='+IntToStr(integer(sgUsers.Objects[0, sgUsers.Row])));
 rgUsers.ItemIndex := StrToInt(fData.cSelectS('RULES', 'TB_USERS', 'where ID='+IntToStr(integer(sgUsers.Objects[0, sgUsers.Row]))));
end;

procedure TfAdmin.FormShow(Sender: TObject);
begin
 Pager.ActivePageIndex := 0;
 PagerChange(self);
 m := 1;
 sgUsersClick(self);
end;

procedure TfAdmin.CreateParams(var Params: TCreateParams);
begin
inherited CreateParams(Params);
with params do
   ExStyle := ExStyle or WS_EX_APPWINDOW;
end;

procedure TfAdmin.bUserEditClick(Sender: TObject);
var r,v: TstringList;
begin
 if ((integer(sgUsers.Objects[0, sgUsers.Row]) = 1)and(rgUsers.ItemIndex = 1)) then
 begin
  fData.ShowError(15);
  exit;
 end;
 if ((Length(Trim(eLogin.Text)) = 0)or(Length(Trim(ePass.Text)) = 0)) then
 begin
  fData.ShowError(11);
  Exit;
 end;
 try
  r := TStringList.Create(); r.Clear();
  v := TStringList.Create(); v.Clear();
  r.Add('LOGIN'); r.Add('PASS'); r.Add('RULES');
  v.Add(''''+Trim(eLogin.Text)+'''');
  v.Add(''''+Trim(ePass.Text)+'''');
  v.Add(IntToStr(rgUsers.ItemIndex));
  fData.cUpdates('TB_USERS', r,v, 'where ID='+IntToStr(integer(sgUsers.Objects[0, sgUsers.Row])));
 finally
  r.Free; v.Free;
  PagerChange(self);
 end;
end;

procedure TfAdmin.bUserAddClick(Sender: TObject);
var r,v: TstringList;
begin
 if ((Length(Trim(eLogin.Text)) = 0)or(Length(Trim(ePass.Text)) = 0)) then
 begin
  fData.ShowError(11);
  Exit;
 end;
 try
  r := TStringList.Create(); r.Clear();
  v := TStringList.Create(); v.Clear();
  r.Add('LOGIN'); r.Add('PASS'); r.Add('RULES');
  v.Add(''''+Trim(eLogin.Text)+'''');
  v.Add(''''+Trim(ePass.Text)+'''');
  v.Add(IntToStr(rgUsers.ItemIndex));
  fData.cInserts('TB_USERS', r,v);
 finally
  r.Free; v.Free;
  PagerChange(self);
 end;
end;

procedure TfAdmin.bUserDelClick(Sender: TObject);
begin
 if (integer(sgUsers.Objects[0, sgUsers.Row]) = 1) then
 begin
  fData.ShowError(16);
  exit;
 end;
 if (MessageDlg(fData.standarts[9], mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
  fData.cDelete('TB_USERS', 'where ID='+IntToStr(integer(sgUsers.Objects[0, sgUsers.Row])));
 PagerChange(self);
end;

procedure TfAdmin.bBDSelectClick(Sender: TObject);
begin
 if (OD.Execute) then
  if (Length(OD.FileName) > 0) then
   eDB.Text := OD.FileName;
end;

procedure TfAdmin.bBDOkClick(Sender: TObject);
var f: TIniFile;
begin
 try
  f := TIniFile.Create(ExtractFileDir(Application.ExeName)+'\config.ini');
  f.WriteString('DataBase', 'Path', Trim(eDB.Text));
  f.WriteString('DataBase', 'user_name', Trim(eDBLogin.Text));
  f.WriteString('DataBase', 'password', Trim(eDBPass.Text));
  f.Free;
  fData.ShowError(17);
  Application.Terminate;
 except
  fData.ShowError(0);
 end;
end;

procedure TfAdmin.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 Action := caFree;
end;

procedure TfAdmin.cbClassInChange(Sender: TObject);
var i: integer;
begin
 cbClassOut.Items.Clear;
 if (cbClassIn.Items.Count = 0) then exit;
 for i := 1 to fData.MaxClass-1 do
  if ((fData.JumpClass[i] = 0)and(fData.CheckJump(i) = 0)and(StrToInt(cbClassIn.Items[cbClassIn.ItemIndex]) <> i)) then
   cbClassOut.Items.Add(IntToStr(i));
 cbClassOut.ItemIndex := 0;
end;

procedure TfAdmin.bClassJAddClick(Sender: TObject);
var f: TIniFile;
    i: integer;
begin
 f := TIniFile.Create(ExtractFileDir(Application.ExeName)+'\config.ini');
 f.WriteString('JumpClass',cbClassIn.Items[cbClassIn.ItemIndex],cbClassOut.Items[cbClassOut.ItemIndex]);
 // Забиваем массив по "перепрыгиванию" классов
 SetLength(fData.JumpClass, fData.MaxClass);
 for i := 0 to fData.MaxClass-1 do
  fData.JumpClass[i] := f.ReadInteger('JumpClass', IntToStr(i), 0);
 f.Free;
 PagerChange(self);
end;

procedure TfAdmin.bClassJDelClick(Sender: TObject);
var f: TIniFile;
    i: integer;
begin
 f := TIniFile.Create(ExtractFileDir(Application.ExeName)+'\config.ini');
 f.DeleteKey('JumpClass',sgJump.Cells[0, sgJump.Row]);
 // Забиваем массив по "перепрыгиванию" классов
 SetLength(fData.JumpClass, fData.MaxClass);
 for i := 0 to fData.MaxClass-1 do
  fData.JumpClass[i] := f.ReadInteger('JumpClass', IntToStr(i), 0);
 f.Free;
 PagerChange(self);
end;

end.
