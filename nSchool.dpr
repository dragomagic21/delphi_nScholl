program nSchool;

uses
  Forms,
  uMain in 'uMain.pas' {fMain},
  uData in 'uData.pas' {fData: TDataModule},
  uPredmet in 'uPredmet.pas' {fPredmet},
  uNewClass in 'uNewClass.pas' {fNewClass},
  uPlanImport in 'uPlanImport.pas' {fPlanImport},
  uClassOb in 'uClassOb.pas' {fClassOb},
  uClassRz in 'uClassRz.pas' {fClassRz},
  uUPlanAdd in 'uUPlanAdd.pas' {fUPlanAdd},
  uSchool in 'uSchool.pas' {fSchool},
  uFind in 'uFind.pas' {fFind},
  uAdmin in 'uAdmin.pas' {fAdmin},
  uLogin in 'uLogin.pas' {fLogin},
  uAbout in 'uAbout.pas' {fAbout};

{$R *.res}

begin
 try
  Application.Initialize;
  Application.HelpFile := 'nSchool.chm';
  Application.Title := 'ÀÐÌ "nSchool"';
  Application.CreateForm(TfData, fData);
  if (fData.Term = 1) then Application.Terminate
  else
  Application.CreateForm(TfMain, fMain);
  Application.Run;
 except
 end;
end.
