unit uAbout;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, AdvOfficePager, StdCtrls, AdvReflectionLabel;

type
  TfAbout = class(TForm)
    Pager: TAdvOfficePager;
    Page: TAdvOfficePage;
    lAuthor: TAdvReflectionLabel;
    lCopyR: TLabel;
    lQuest: TLabel;
    lMail: TLabel;
    procedure PagerClosePage(Sender: TObject; PageIndex: Integer;
      var Allow: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure lMailClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  fAbout: TfAbout;

implementation

uses uData, ShellAPI;

{$R *.dfm}

procedure TfAbout.PagerClosePage(Sender: TObject; PageIndex: Integer; var Allow: Boolean);
begin
 Self.Close();
end;

procedure TfAbout.FormCreate(Sender: TObject);
begin
 fData.SetLanguage(Self);
end;

procedure TfAbout.lMailClick(Sender: TObject);
begin
 ShellExecute(Handle,'open','mailto:drago_magic@mail.ru',nil,nil,SW_SHOWNORMAL);
end;

end.
