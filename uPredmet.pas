unit uPredmet;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, AdvEdit, AdvGlassButton, Grids, BaseGrid, AdvGrid,
  ExtCtrls, AdvPanel, AdvObj;

type
  TfPredmet = class(TForm)
    Panel: TAdvPanel;
    sgPr: TAdvStringGrid;
    bPrAdd: TAdvGlassButton;
    bPrEdit: TAdvGlassButton;
    bPrDel: TAdvGlassButton;
    ePrName: TAdvEdit;
    LName: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bPrAddClick(Sender: TObject);
    procedure bPrEditClick(Sender: TObject);
    procedure bPrDelClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    procedure CreateParams(var Params: TCreateParams); override;
  public
    { Public declarations }
  end;

var
  fPredmet: TfPredmet;

implementation

uses uData;

{$R *.dfm}

procedure TfPredmet.FormCreate(Sender: TObject);
begin
 // Устанавливаем язык для текущей формы
 fData.SetLanguage(fPredmet);
end;

procedure TfPredmet.FormShow(Sender: TObject);
begin
 // Выводим список предметов
 case fData.cFillSg(sgPr,8,'NAME','TB_PREDMET') of
  // Врубаем кнопки Редактирования и Удаления
  0: begin
      bPrEdit.Enabled := TRUE;
      bPrDel.Enabled  := TRUE;
     end;
  // Вырубаем кнопки Редактирования и Удаления
  1: begin
      bPrEdit.Enabled := FALSE;
      bPrDel.Enabled  := FALSE;
     end;
 end;
end;

procedure TfPredmet.bPrAddClick(Sender: TObject);
begin
 case fData.cInsert('TB_PREDMET','NAME',''''+Trim(ePrName.Text)+'''') of
  0: FormShow(self);
 end;
end;

procedure TfPredmet.bPrEditClick(Sender: TObject);
begin
 case fData.cUpdate('TB_PREDMET','NAME',''''+Trim(ePrName.Text)+'''','where ID='+IntToStr(integer(sgPr.Objects[0, sgPr.Row]))) of
  0:  FormShow(self);
 end;
end;

procedure TfPredmet.bPrDelClick(Sender: TObject);
begin
 // Проверяем наличие предмета в учебных планах - если есть - нельзя удалять
 if fData.cCount('ID','TB_UPLAN','WHERE PREDMET='+IntToStr(integer(sgPr.Objects[0, sgPr.Row]))) > 0 then
  fData.ShowError(9,Self.Handle) else
 // Проверяем уверенность юзера в своих намерениях =)
 if (MessageDlg(fData.standarts[9],mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
 case fData.cDelete('TB_PREDMET','where ID='+IntToStr(integer(sgPr.Objects[0, sgPr.Row]))) of
  0:  FormShow(self);
 end;
end;

procedure TfPredmet.CreateParams(var Params: TCreateParams);
begin
inherited CreateParams(Params);
with params do
   ExStyle := ExStyle or WS_EX_APPWINDOW;
end;

procedure TfPredmet.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 Action := caFree;
end;

end.
