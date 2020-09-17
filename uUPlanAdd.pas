unit uUPlanAdd;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, AdvFontCombo, StdCtrls, AdvEdit, AdvGlassButton, ExtCtrls,
  AdvPanel;

type
  TfUPlanAdd = class(TForm)
    Panel: TAdvPanel;
    LName: TLabel;
    bCancel: TAdvGlassButton;
    bOk: TAdvGlassButton;
    EName: TAdvEdit;
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
  fUPlanAdd: TfUPlanAdd;

implementation

uses uData, uMain;

{$R *.dfm}

{ TfUPlanAdd }

procedure TfUPlanAdd.CreateParams(var Params: TCreateParams);
begin
inherited CreateParams(Params);
with params do
   ExStyle := ExStyle or WS_EX_APPWINDOW;
end;

procedure TfUPlanAdd.bOkClick(Sender: TObject);
begin
 case Mode of
 // ДОБАВЛЕНИЕ НОВОГО КЛАССА
  0: try
      case fData.cInsert('TB_UPLAN','SCHOOL,NAME',IntToStr(School)+','+''''+Trim(EName.Text)+'''') of
       0: begin
           fMain.PagerChange(self);
           Self.Close();
          end;
       1: fData.ShowError(6,Self.Handle);
      end;
     except
      fData.ShowError(1,Self.Handle);
     end;
 // ОБНОВЛЕНИЕ ИНФОРМАЦИИ
  1: begin
      try
         case fData.cUpdate('TB_UPLAN','NAME',''''+Trim(EName.Text)+'''','WHERE ID='+IntToStr(ID)) of
          0: begin
              fMain.PagerChange(self);
              Self.Close();
             end;
          1: fData.ShowError(6,Self.Handle);
         end;
       except
        fData.ShowError(1,Self.Handle);
       end;
     end;
 end;
end;

procedure TfUPlanAdd.FormCreate(Sender: TObject);
begin
 fData.SetLanguage(Self);
end;

procedure TfUPlanAdd.FormShow(Sender: TObject);
begin
 // Вычисляем заголовок окна
  case Mode of
   0: Self.Caption := fData.standarts[2]+' '+fData.standarts[24]+' '+fData.cSelectS('NAME','TB_SCHOOL','WHERE ID='+IntToStr(School));
   1: Self.Caption := fData.standarts[11]+' '+fData.standarts[24]+' '+fData.cSelectS('NAME','TB_SCHOOL','WHERE ID='+IntToStr(School));
  end;
 // ЕСЛИ РЕДАКТИРОВАНИЕ - ЗАПОЛНЯЕМ ПОЛЯ
 if (Mode = 1) then
  EName.Text := fData.cSelectS('NAME','TB_UPLAN','WHERE ID='+IntToStr(ID));
end;

procedure TfUPlanAdd.bCancelClick(Sender: TObject);
begin
 Self.Close();
end;

procedure TfUPlanAdd.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 Action := caFree;
end;

end.
