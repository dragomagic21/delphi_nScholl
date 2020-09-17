unit uSchool;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtDlgs, AdvGlassButton, ExtCtrls, StdCtrls, AdvEdit,
  AdvOfficePager;

type
  TfSchool = class(TForm)
    Pager: TAdvOfficePager;
    Page: TAdvOfficePage;
    lName: TLabel;
    eName: TAdvEdit;
    lGerb: TLabel;
    bFile: TAdvGlassButton;
    bClear: TAdvGlassButton;
    bSave: TAdvGlassButton;
    opd: TOpenPictureDialog;
    Image1: TImage;
    procedure bFileClick(Sender: TObject);
    procedure bClearClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bSaveClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Private declarations }
    procedure CreateParams(var Params: TCreateParams); override;
  public
    { Public declarations }
    Mode: byte;
    ID: integer;
  end;

var
  fSchool: TfSchool;

implementation

uses uData, uMain;

{$R *.dfm}

procedure TfSchool.bFileClick(Sender: TObject);
begin
 if (opd.Execute) then
  if (Length(opd.FileName) > 0) then
   Image1.Picture.LoadFromFile(opd.FileName);
 Image1.Refresh();
end;

procedure TfSchool.bClearClick(Sender: TObject);
begin
 if (Mode = 1) then
 begin
  fData.cUpdate('TB_SCHOOL','IMG','1','where ID='+IntToStr(ID));
  if FileExists(ExtractFileDir(Application.ExeName)+'\imgs\'+IntToStr(ID)+'.jpg') then
     DeleteFile(ExtractFileDir(Application.ExeName)+'\imgs\'+IntToStr(ID)+'.jpg');
  Image1.Picture := nil;
  Image1.Picture.Graphic := nil;
 end;
end;

procedure TfSchool.CreateParams(var Params: TCreateParams);
begin
inherited CreateParams(Params);
with params do
   ExStyle := ExStyle or WS_EX_APPWINDOW;
end;

procedure TfSchool.FormShow(Sender: TObject);
begin
 if (Mode = 1) then
 begin
  eName.Text := fData.cSelectS('NAME','TB_SCHOOL','where ID='+IntToStr(ID));
  fData.FillImg(Image1, ID);
 end;
end;

procedure TfSchool.bSaveClick(Sender: TObject);
begin
 if (Length(Trim(eName.Text)) = 0) then fData.ShowError(11) else
 case Mode of
  0: begin
      fData.cInsert('TB_SCHOOL', 'NAME', ''''+Trim(eName.Text)+'''');
      if (Image1.Picture.Graphic = nil) then fData.cUpdate('TB_SCHOOL','IMG','0','where ID='+fData.cSelectS('max(ID)','TB_SCHOOL')) else
      begin
       fData.cUpdate('TB_SCHOOL','IMG','1','where ID='+fData.cSelectS('max(ID)','TB_SCHOOL'));
       Image1.Picture.SaveToFile(ExtractFileDir(fData.Database.DatabaseName)+'\imgs\'+fData.cSelectS('max(ID)','TB_SCHOOL')+'.jpg');
      end;
     end;
  1: begin
      fData.cUpdate('TB_SCHOOL', 'NAME', ''''+Trim(eName.Text)+'''', 'where ID='+IntToStr(ID));
      if (Image1.Picture.Graphic = nil) then fData.cUpdate('TB_SCHOOL','IMG','0','where ID='+IntToStr(ID)) else
      begin
       fData.cUpdate('TB_SCHOOL','IMG','1','where ID='+IntToStr(ID));
       Image1.Picture.SaveToFile(ExtractFileDir(fData.Database.DatabaseName)+'\imgs\'+IntToStr(ID)+'.jpg');
      end;
     end;
 end;
 fMain.OnShow(Self);
 Self.Close();
end;

procedure TfSchool.FormCreate(Sender: TObject);
begin
 fData.SetLanguage(Self);
 Image1.Picture := nil
end;

procedure TfSchool.FormClose(Sender: TObject; var Action: TCloseAction);
begin
 Action := caFree;
end;

end.
