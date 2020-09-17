unit uNewClass;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, AdvGlassButton, StdCtrls, AdvEdit, ExtCtrls, AdvPanel,
  AdvFontCombo;

type
  TfNewClass = class(TForm)
    Panel: TAdvPanel;
    bCancel: TAdvGlassButton;
    bOk: TAdvGlassButton;
    EName: TAdvEdit;
    LName: TLabel;
    cbNum: TAdvOfficeComboBox;
    Label1: TLabel;
    cbUPlan: TAdvOfficeComboBox;
    lClassPlan: TLabel;
    procedure bOkClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bCancelClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    procedure CreateParams(var Params: TCreateParams); override;
  public
    Mode: byte;      // Тип открытия формы - добавление/редактирование класса
    ID: integer;     // Используется при редактировании класса
    School: integer; // Ссылка на школу, к которой относиться класс
    { Public declarations }
  end;

var
  fNewClass: TfNewClass;

implementation

uses uData, uMain;

{$R *.dfm}

{ TfNewClass }

procedure TfNewClass.CreateParams(var Params: TCreateParams);
begin
inherited CreateParams(Params);
with params do
   ExStyle := ExStyle or WS_EX_APPWINDOW;
end;

procedure TfNewClass.bOkClick(Sender: TObject);
var log: string;
    r,v: TStringList;
begin
 case Mode of
 // ДОБАВЛЕНИЕ НОВОГО КЛАССА
  0: try
      log := (DateToStr(Now())+' '+fData.standarts[3]+' '+fData.standarts[4]+' '+fData.cSelectS('NAME','TB_SCHOOL','WHERE ID='+IntToStr(School)));
      case fData.cInsert('TB_CLASS','SCHOOL,NUM,NAME,LOG',IntToStr(School)+','+cbNum.Items.Strings[cbNum.ItemIndex]+','+''''+Trim(EName.Text)+''''+','+''''+log+'''') of
       0: begin
           fMain.sgSchoolsClick(self);
           Self.Close();
          end;
       1: fData.ShowError(6,Self.Handle);
      end;
     except
      fData.ShowError(1,Self.Handle);
     end;
 // ОБНОВЛЕНИЕ ИНФОРМАЦИИ
  1: begin
      r := TStringList.Create;
      v := TStringList.Create;
      try
       log := fData.cSelectS('LOG','TB_CLASS','WHERE ID='+IntToStr(ID));
       try
        r.Add('NUM'); r.Add('NAME'); r.Add('LOG');
        v.Add(cbNum.Items.Strings[cbNum.ItemIndex]);
        v.Add(''''+Trim(EName.Text)+'''');
        v.Add(''''+log+'#10#13'+DateToStr(Now())+' '+fData.standarts[5]+'''');
         case fData.cUpdates('TB_CLASS',r,v,'WHERE ID='+IntToStr(ID)) of
          0: begin
              fMain.sgSchoolsClick(self);
              Self.Close();
             end;
          1: fData.ShowError(6,Self.Handle);
         end;
       except
        fData.ShowError(1,Self.Handle);
       end;
      finally
       r.Free; v.Free;
      end;
     end;
 end;
end;

procedure TfNewClass.FormCreate(Sender: TObject);
var i: integer;
begin
 fData.SetLanguage(Self);
 cbNum.Items.Clear();
 for i:=1 to fData.MaxClass-1 do
  cbNum.Items.Add(IntToStr(i));
 cbUPlan.Items.Clear();
end;

procedure TfNewClass.FormShow(Sender: TObject);
begin
 // Вычисляем заголовок окна
  case Mode of
   0: Self.Caption := fData.standarts[2]+' '+fData.standarts[10]+' '+fData.cSelectS('NAME','TB_SCHOOL','WHERE ID='+IntToStr(School));
   1: Self.Caption := fData.standarts[11]+' '+fData.standarts[10]+' '+fData.cSelectS('NAME','TB_SCHOOL','WHERE ID='+IntToStr(School));
  end;
 // ЕСЛИ РЕДАКТИРОВАНИЕ - ЗАПОЛНЯЕМ ПОЛЯ
 if (Mode = 1) then
 begin
  cbNum.ItemIndex := cbNum.Items.IndexOf(fData.cSelectS('NUM','TB_CLASS','WHERE ID='+IntToStr(ID)));
  EName.Text := fData.cSelectS('NAME','TB_CLASS','WHERE ID='+IntToStr(ID));
 end;
 fData.cFillCb(cbUPlan,'NAME','TB_UPLAN','Where (SCHOOL='+IntToStr(School)+')');
 if (Mode = 1) then
  cbUPlan.Items.IndexOf(fData.cSelectS('NAME','TB_UPLAN','where (ID=(select UPLAN from TB_CLASS where ID='+IntToStr(ID)+'))'));
end;

procedure TfNewClass.bCancelClick(Sender: TObject);
begin
 Self.Close();
end;

procedure TfNewClass.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 Action := caFree;
end;

end.
