unit uMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, AdvOfficePager, ExtCtrls, AdvPanel, Grids, BaseGrid, AdvGrid,
  ComCtrls, AdvOfficeTabSet, AdvGlowButton, AdvToolBtn, AdvToolBar, Menus,
  AdvMenus, StdCtrls, AdvCombo, AdvGroupBox, AdvOfficeButtons,
  AdvReflectionLabel, AdvFontCombo, AdvGlassButton, AdvEdit, AdvPicture, AdvObj,
  AdvSplitter;

type
  TfMain = class(TForm)
    Pager: TAdvOfficePager;
    pClass: TAdvOfficePage;
    pTeacher: TAdvOfficePage;
    pPlans: TAdvOfficePage;
    pNav: TAdvPanel;
    pClassInfo: TAdvOfficePager;
    tClassInfoPeop: TAdvOfficePage;
    tClassInfoMarks: TAdvOfficePage;
    tClassInfoMed: TAdvOfficePage;
    tClassInfoPeopNav: TAdvPanel;
    tClassInfoPeopAll: TAdvPanel;
    tClassInfoPeopAbout: TAdvPanel;
    tClassInfoMarksNav: TAdvPanel;
    tClassInfoMarksPSg: TAdvPanel;
    sgClassInfoMarks: TAdvStringGrid;
    sgClassInfoMarksAll: TAdvStringGrid;
    tTeacherNav: TAdvPanel;
    tTeacherAll: TAdvPanel;
    tTeacherInfo: TAdvPanel;
    tPlansClass: TAdvPanel;
    tPlansSg: TAdvPanel;
    mMenu: TAdvMainMenu;
    tInfoNav: TAdvPanel;
    MnHelp: TMenuItem;
    MnHelpInfo: TMenuItem;
    MnBreak1: TMenuItem;
    MnAbout: TMenuItem;
    LSchoolName: TAdvReflectionLabel;
    sgClassInfo: TAdvStringGrid;
    bClassAdd: TAdvGlassButton;
    bClassEdit: TAdvGlassButton;
    bClassDel: TAdvGlassButton;
    bPeopAdd: TAdvGlassButton;
    bPeopDel: TAdvGlassButton;
    EPeopFam: TAdvEdit;
    LPeopFam: TLabel;
    EPeopName: TAdvEdit;
    LPeopName: TLabel;
    EPeopSname: TAdvEdit;
    LPeopSname: TLabel;
    bSavePeopInfo: TAdvGlassButton;
    Bevel1: TBevel;
    LPeopMoveClass: TLabel;
    cbPeopMoveClass: TAdvOfficeComboBox;
    bPeopMoveClass: TAdvGlassButton;
    Bevel2: TBevel;
    LPeopMoveSchool: TLabel;
    cbPeopMoveSchool: TAdvOfficeComboBox;
    bPeopMoveSchool: TAdvGlassButton;
    cbPeopMoveSchoolClass: TAdvOfficeComboBox;
    bTeacherAdd: TAdvGlassButton;
    bTeacherDel: TAdvGlassButton;
    sgTeacher: TAdvStringGrid;
    bTeacherSave: TAdvGlassButton;
    LTeacherMoveSchool: TLabel;
    cbTeacherMoveSchool: TAdvOfficeComboBox;
    bTeacherMoveSchool: TAdvGlassButton;
    bPredmet: TAdvGlassButton;
    MnFind: TMenuItem;
    PagerPlans: TAdvOfficePager;
    PPlansOpr: TAdvOfficePage;
    LPlanTeacher: TLabel;
    LPlanPredmet: TLabel;
    cbPlansTeach: TAdvOfficeComboBox;
    cbPlansPr: TAdvOfficeComboBox;
    bPlanRAdd: TAdvGlassButton;
    bPlanRDel: TAdvGlassButton;
    bPlanImport: TAdvGlassButton;
    PPlansFilter: TAdvOfficePage;
    LPPlansFilterPr: TLabel;
    LPPlansFilterTch: TLabel;
    Bevel3: TBevel;
    bPPlansFilterPr: TAdvGlassButton;
    bPPlansFilterTch: TAdvGlassButton;
    sgPlansClass: TAdvStringGrid;
    ePPlansFilterPr: TAdvOfficeComboBox;
    ePPlansFilterTch: TAdvOfficeComboBox;
    pClassInfoMarks: TAdvPanel;
    rClassInfoMarksP: TAdvOfficeRadioButton;
    lClassInfoMarks: TLabel;
    rClassInfoMarksW: TAdvOfficeRadioButton;
    Bevel4: TBevel;
    bMarksSave: TAdvGlassButton;
    bMarksExcel: TAdvGlassButton;
    bMarksChart: TAdvGlassButton;
    tPlansClassName: TAdvOfficeTabSet;
    pUplanNav: TAdvPanel;
    bPlanAdd: TAdvGlassButton;
    bPlanEdit: TAdvGlassButton;
    bPlanDel: TAdvGlassButton;
    pEnd: TAdvOfficePage;
    pClassMed: TAdvPanel;
    sgClassMedPeop: TAdvStringGrid;
    pClassMedNav: TAdvPanel;
    pClassMedAdd: TAdvPanel;
    pClassMedFilter: TAdvPanel;
    sgClassMed: TAdvStringGrid;
    eClassMedDB: TDateTimePicker;
    eClassMedDE: TDateTimePicker;
    eClassMedTxt: TAdvEdit;
    bClassMedAdd: TAdvGlassButton;
    lClassMedDB: TLabel;
    lClassMedDE: TLabel;
    lClassMedTxt: TLabel;
    cClassMedFilter: TAdvOfficeCheckBox;
    eClassMedFilter: TDateTimePicker;
    tSchools: TAdvOfficeTabSet;
    pPredmet: TAdvPanel;
    pAbout: TAdvOfficePage;
    pClassName: TAdvPanel;
    sgClassNames: TAdvStringGrid;
    pClassNav: TAdvPanel;
    tClassAbout: TAdvOfficePage;
    lClassPlan: TLabel;
    cbClassPlan: TAdvOfficeComboBox;
    bClassPlan: TAdvGlassButton;
    lClassPeopCount: TLabel;
    lClassPeopCountR: TLabel;
    pAMarks: TAdvOfficePage;
    pAMarksNav: TAdvPanel;
    pAMarksExcel: TAdvGlassButton;
    pAMarksChart: TAdvGlassButton;
    pAMarksPeop: TAdvPanel;
    sgAMarksPeop: TAdvStringGrid;
    pAMarksM: TAdvPanel;
    sgAMarksM: TAdvStringGrid;
    lClassAMarksClass: TLabel;
    cClassAMarksClass: TAdvOfficeComboBox;
    Bevel5: TBevel;
    bClassAMarksClass: TAdvGlassButton;
    pEndNav: TAdvOfficePager;
    pEndInfo: TAdvOfficePage;
    lEndInfoCount: TLabel;
    lEndInfoCountR: TLabel;
    pEndInfoNav: TAdvOfficePage;
    pEndInfoPeop: TAdvPanel;
    sgEndInfoPeop: TAdvStringGrid;
    AdvPanel5: TAdvPanel;
    lpEndInfoFName: TLabel;
    lpEndInfoName: TLabel;
    lpEndInfoSName: TLabel;
    epEndInfoFName: TAdvEdit;
    epEndInfoName: TAdvEdit;
    epEndInfoSName: TAdvEdit;
    pEndMed: TAdvOfficePage;
    pEndMedPeop: TAdvPanel;
    sgEndMedPeop: TAdvStringGrid;
    pEndMedp: TAdvPanel;
    pEndMedFilter: TAdvPanel;
    cEndMedFilter: TAdvOfficeCheckBox;
    eEndMedFilter: TDateTimePicker;
    sgEndMed: TAdvStringGrid;
    pEndAMarks: TAdvOfficePage;
    pEndAMarksNav: TAdvPanel;
    pEndAMarksClass: TLabel;
    Bevel9: TBevel;
    bEndAMarksExcel: TAdvGlassButton;
    bEndAMarksChart: TAdvGlassButton;
    eEndAMarksClass: TAdvOfficeComboBox;
    bEndAMarksClass: TAdvGlassButton;
    pEndAMarksPeop: TAdvPanel;
    sgEndAMarksPeop: TAdvStringGrid;
    pEndAMarksM: TAdvPanel;
    sgEndAMarksM: TAdvStringGrid;
    pEndClasses: TAdvPanel;
    sgEndClasses: TAdvStringGrid;
    lCountMarks: TLabel;
    lM1012: TLabel;
    lM79: TLabel;
    lM46: TLabel;
    lM13: TLabel;
    lM1012c: TLabel;
    lM79c: TLabel;
    lM46c: TLabel;
    lM13c: TLabel;
    lSrClassMark: TLabel;
    lCSrClassMark: TLabel;
    Bevel6: TBevel;
    lSPeop: TLabel;
    lCSPeop: TLabel;
    lSrSMark: TLabel;
    lCSrSMark: TLabel;
    lSMarks: TLabel;
    lSMarks10: TLabel;
    lSMarks7: TLabel;
    lSMarks4: TLabel;
    lSMarks1: TLabel;
    lCSMarks1: TLabel;
    lCSMarks4: TLabel;
    lCSMarks7: TLabel;
    lCSMarks10: TLabel;
    lSTeacher: TLabel;
    lCSTeacher: TLabel;
    lSParalel: TLabel;
    sgSchoolStat: TAdvStringGrid;
    Bevel7: TBevel;
    Bevel8: TBevel;
    Bevel10: TBevel;
    Bevel11: TBevel;
    Bevel12: TBevel;
    Image1: TImage;
    bClassIn: TAdvGlassButton;
    bClassBreak: TAdvGlassButton;
    bSchoolAdd: TAdvGlassButton;
    bSchoolEdit: TAdvGlassButton;
    bSchoolDel: TAdvGlassButton;
    Bevel13: TBevel;
    pmSchool: TAdvPopupMenu;
    pmSchoolAdd: TMenuItem;
    pmSchoolEdit: TMenuItem;
    pmSchoolDel: TMenuItem;
    pmClass: TAdvPopupMenu;
    pmClassAdd: TMenuItem;
    pmClassEdit: TMenuItem;
    pmClassDel: TMenuItem;
    pmClass1: TMenuItem;
    pmClassIn: TMenuItem;
    pmClassBreak: TMenuItem;
    Bevel14: TBevel;
    bNextYear: TAdvGlassButton;
    bMedExcel: TAdvGlassButton;
    bEndAMedExcel: TAdvGlassButton;
    MnAdmin: TMenuItem;
    MnLanguage: TMenuItem;
    ODLang: TOpenDialog;
    ePeopBirthDay: TDateTimePicker;
    LPeopBirthDay: TLabel;
    ePeopAdr: TAdvEdit;
    LPeopAdr: TLabel;
    LPeopP: TLabel;
    LPeopINN: TLabel;
    LPeopPrim: TLabel;
    ePeopPS: TAdvEdit;
    ePeopPN: TAdvEdit;
    ePeopINN: TAdvEdit;
    ePeopPrim: TMemo;
    LTeacherFname: TLabel;
    eTeacherFname: TAdvEdit;
    eTeacherName: TAdvEdit;
    LTeacherName: TLabel;
    eTeacherSname: TAdvEdit;
    LTeacherSname: TLabel;
    eTeacherBirthDay: TDateTimePicker;
    LTeacherBirthDay: TLabel;
    LTeacherAdr: TLabel;
    eTeacherAdr: TAdvEdit;
    LTeacherP: TLabel;
    eTeacherPS: TAdvEdit;
    eTeacherPN: TAdvEdit;
    eTeacherINN: TAdvEdit;
    LTeacherINN: TLabel;
    eTeacherPrim: TMemo;
    LTeacherPrim: TLabel;
    Bevel15: TBevel;
    bPeopWord: TAdvGlassButton;
    bTeacherWord: TAdvGlassButton;
    bPPlansReportExcel: TAdvGlassButton;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    AdvEdit1: TAdvEdit;
    Label13: TLabel;
    AdvEdit2: TAdvEdit;
    AdvEdit3: TAdvEdit;
    Label14: TLabel;
    AdvEdit4: TAdvEdit;
    Label15: TLabel;
    Memo1: TMemo;
    DateTimePicker1: TDateTimePicker;
    Label16: TLabel;
    AdvGlassButton1: TAdvGlassButton;
    AdvSplitter1: TAdvSplitter;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure tSchoolsChange(Sender: TObject);
    procedure tClassChange(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure sgSchoolsClick(Sender: TObject);
    procedure pClassInfoChange(Sender: TObject);
    procedure sgClassInfoClick(Sender: TObject);
    procedure PagerChange(Sender: TObject);
    procedure sgTeacherClick(Sender: TObject);
    procedure bPeopAddClick(Sender: TObject);
    procedure bSavePeopInfoClick(Sender: TObject);
    procedure bTeacherAddClick(Sender: TObject);
    procedure bTeacherSaveClick(Sender: TObject);
    procedure tPlansClassNameChange(Sender: TObject);
    procedure bPlanRAddClick(Sender: TObject);
    procedure bPredmetClick(Sender: TObject);
    procedure bPlanRDelClick(Sender: TObject);
    procedure bTeacherDelClick(Sender: TObject);
    procedure bPeopDelClick(Sender: TObject);
    procedure bClassAddClick(Sender: TObject);
    procedure bClassEditClick(Sender: TObject);
    procedure PagerPlansChange(Sender: TObject);
    procedure bPPlansFilterPrClick(Sender: TObject);
    procedure bPPlansFilterTchClick(Sender: TObject);
    procedure bPPlansReportExcelClick(Sender: TObject);
    procedure bPlanImportClick(Sender: TObject);
    procedure bClassDelClick(Sender: TObject);
    procedure rClassInfoMarksPClick(Sender: TObject);
    procedure rClassInfoMarksWClick(Sender: TObject);
    procedure sgClassInfoMarksAllClick(Sender: TObject);
    procedure sgClassInfoMarksClick(Sender: TObject);
    procedure bMarksSaveClick(Sender: TObject);
    procedure bMarksExcelClick(Sender: TObject);
    procedure bMarksChartClick(Sender: TObject);
    procedure MnClassInClick(Sender: TObject);
    procedure MnClassOutClick(Sender: TObject);
    procedure bPlanAddClick(Sender: TObject);
    procedure bPlanEditClick(Sender: TObject);
    procedure sgClassMedPeopClick(Sender: TObject);
    procedure tClassInfoMedShow(Sender: TObject);
    procedure cClassMedFilterClick(Sender: TObject);
    procedure bClassMedAddClick(Sender: TObject);
    procedure sgClassNamesClick(Sender: TObject);
    procedure tSchoolsTabClick(Sender: TObject; PageIndex: Integer);
    procedure bClassPlanClick(Sender: TObject);
    procedure sgAMarksPeopClick(Sender: TObject);
    procedure bClassAMarksClassClick(Sender: TObject);
    procedure pAMarksExcelClick(Sender: TObject);
    procedure pAMarksChartClick(Sender: TObject);
    procedure sgEndClassesClick(Sender: TObject);
    procedure pEndNavChange(Sender: TObject);
    procedure sgEndInfoPeopClick(Sender: TObject);
    procedure sgEndAMarksPeopClick(Sender: TObject);
    procedure sgEndMedPeopClick(Sender: TObject);
    procedure bEndAMarksClassClick(Sender: TObject);
    procedure bEndAMarksExcelClick(Sender: TObject);
    procedure bEndAMarksChartClick(Sender: TObject);
    procedure sgClassInfoMarksKeyPress(Sender: TObject; var Key: Char);
    procedure MnSchoolAddClick(Sender: TObject);
    procedure MnSchoolEditClick(Sender: TObject);
    procedure MnSchoolDelClick(Sender: TObject);
    procedure MnClassAddClick(Sender: TObject);
    procedure MnClassEditClick(Sender: TObject);
    procedure MnClassDelClick(Sender: TObject);
    procedure bClassInClick(Sender: TObject);
    procedure bClassBreakClick(Sender: TObject);
    procedure bSchoolAddClick(Sender: TObject);
    procedure bSchoolEditClick(Sender: TObject);
    procedure bSchoolDelClick(Sender: TObject);
    procedure pmSchoolAddClick(Sender: TObject);
    procedure pmSchoolEditClick(Sender: TObject);
    procedure pmSchoolDelClick(Sender: TObject);
    procedure pmClassAddClick(Sender: TObject);
    procedure pmClassEditClick(Sender: TObject);
    procedure pmClassDelClick(Sender: TObject);
    procedure pmClassInClick(Sender: TObject);
    procedure pmClassBreakClick(Sender: TObject);
    procedure MnFindClick(Sender: TObject);
    procedure bNextYearClick(Sender: TObject);
    procedure bPeopMoveClassClick(Sender: TObject);
    procedure cbPeopMoveSchoolChange(Sender: TObject);
    procedure bPeopMoveSchoolClick(Sender: TObject);
    procedure bMedExcelClick(Sender: TObject);
    procedure bEndAMedExcelClick(Sender: TObject);
    procedure MnAdminClick(Sender: TObject);
    procedure MnLanguageClick(Sender: TObject);
    procedure bTeacherMoveSchoolClick(Sender: TObject);
    procedure bPeopWordClick(Sender: TObject);
    procedure bTeacherWordClick(Sender: TObject);
    procedure MnExitClick(Sender: TObject);
    procedure MnAboutClick(Sender: TObject);
    procedure MnHelpInfoClick(Sender: TObject);
    procedure bPlanDelClick(Sender: TObject);
    procedure AdvGlassButton1Click(Sender: TObject);
    procedure AdvSplitter1Moved(Sender: TObject);
  private
    UPlanFilter: byte;
    UPlanFilterID: integer;
    { Private declarations }
  public
    Mode: byte;
    { Public declarations }
    procedure SetRules(Filter: byte = 0);
  end;

var
  fMain: TfMain;

implementation

uses uData, uPredmet, uNewClass, uPlanImport, uClassOb, uClassRz, uUPlanAdd,
  uSchool, uFind, uAdmin, uLogin, uAbout, ShellAPI;

{$R *.dfm}

procedure TfMain.FormCreate(Sender: TObject);
begin
 UPlanFilter := 0;
 UPlanFilterID := -1;
 // устанавливаем языковые настройки для текущей формы
 fData.SetLanguage(Self);
 // устанавливаем начальное положение вкладок
 Pager.ActivePageIndex := 0;
 pClassInfo.ActivePageIndex := 0;
 pEndNav.ActivePageIndex := 0;
 // Устанавливаем везде текущую дату
 eClassMedFilter.DateTime := Now();
 eClassMedDB.DateTime := Now();
 eClassMedDE.DateTime := Now();
 eEndMedFilter.DateTime := Now();
 ODLang.InitialDir := ExtractFileDir(Application.ExeName)+'\language\';
 Mode := 0;
end;

procedure TfMain.FormShow(Sender: TObject);
begin
if (fData.Term = 1) then Application.Terminate;
if (Mode = 0) then
begin
 Application.CreateForm(TfLogin, fLogin);
 fLogin.ShowModal();
end;
SetRules();
 // выбираем школы из БД
 case fData.FillSchollsTab(tSchools) of
  // выбока ок
  0: begin
      tSchoolsChange(self);
      // Врубаем кнопки Редактирования и Удаления
//      bSchoolEdit.Enabled := TRUE;
//      bSchoolDel.Enabled  := TRUE;
     end;
  // ошибка или пусто
  1: begin
      tSchoolsChange(self);
      // Вырубаем кнопки Редактирования и Удаления
//      bSchoolEdit.Enabled := FALSE;
//      bSchoolDel.Enabled  := FALSE;
     end;
 end;
end;

procedure TfMain.tSchoolsChange(Sender: TObject);
begin
 PagerChange(Self);
// выбираем перечень классов для текущей школы
// case fData.FillClassTab(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag, tClass) of
//  0: tClassChange(self);
//  1:
// end;
end;

procedure TfMain.tClassChange(Sender: TObject);
begin
 // выбираем список чеников класса
// case fData.FillPeopSg(tClass.AdvOfficeTabs[tClass.ActiveTabIndex].Tag, sgClass) of
//  0: //
//  1:
// end;
end;

procedure TfMain.FormResize(Sender: TObject);
begin
 PagerChange(self);
 cbPeopMoveClass.Top := LPeopMoveClass.Top + 17;
 cbPeopMoveSchool.Top := LPeopMoveClass.Top + 17;
 cbPeopMoveSchoolClass.Top := LPeopMoveClass.Top + 17;
 cbTeacherMoveSchool.Top := LTeacherMoveSchool.Top + 24;
end;

procedure TfMain.sgSchoolsClick(Sender: TObject);
begin
 PagerChange(Self);
end;

procedure TfMain.pClassInfoChange(Sender: TObject);
begin
 // В зависимости от открытой вкладки - обновляем инфу
 case pClassInfo.ActivePageIndex of
  0: begin
      // Учебный план
       if (fData.cFillCb(cbClassPlan, 'NAME', 'TB_UPLAN', 'where (SCHOOL='+IntToStr(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag)+')') = 0) then
        if (fData.cSelectS('UPLAN','TB_CLASS','where ID='+IntToStr(integer(sgClassNames.Objects[0, sgClassNames.Row]))) <> '') then
         cbClassPlan.ItemIndex := cbClassPlan.Items.IndexOfObject(Pointer(StrToInt(fData.cSelectS('UPLAN','TB_CLASS','where ID='+IntToStr(integer(sgClassNames.Objects[0, sgClassNames.Row])))))) else
         cbClassPlan.ItemIndex := -1;
      // Количество учеников
      lClassPeopCountR.Caption := IntToStr(fData.cCount('ID','TB_PEOP','where CLASS='+IntToStr(integer(sgClassNames.Objects[0, sgClassNames.Row]))));
      // Распределение учеников по баллам
      lM1012c.Caption := IntToStr(fData.cCount('TB_PEOP.ID','TB_PEOP, TB_MARKS','where ((TB_PEOP.ID=TB_MARKS.PEOP)and(TB_PEOP.CLASS='+IntToStr(integer(sgClassNames.Objects[0, sgClassNames.Row]))+')and(TB_MARKS.FL=0)and(cast(TB_MARKS.YER as integer) between 10 and 12))'));
      lM79c.Caption := IntToStr(fData.cCount('TB_PEOP.ID','TB_PEOP, TB_MARKS','where ((TB_PEOP.ID=TB_MARKS.PEOP)and(TB_PEOP.CLASS='+IntToStr(integer(sgClassNames.Objects[0, sgClassNames.Row]))+')and(TB_MARKS.FL=0)and(cast(TB_MARKS.YER as integer) between 7 and 9))'));
      lM46c.Caption := IntToStr(fData.cCount('TB_PEOP.ID','TB_PEOP, TB_MARKS','where ((TB_PEOP.ID=TB_MARKS.PEOP)and(TB_PEOP.CLASS='+IntToStr(integer(sgClassNames.Objects[0, sgClassNames.Row]))+')and(TB_MARKS.FL=0)and(cast(TB_MARKS.YER as integer) between 4 and 6))'));
      lM13c.Caption := IntToStr(fData.cCount('TB_PEOP.ID','TB_PEOP, TB_MARKS','where ((TB_PEOP.ID=TB_MARKS.PEOP)and(TB_PEOP.CLASS='+IntToStr(integer(sgClassNames.Objects[0, sgClassNames.Row]))+')and(TB_MARKS.FL=0)and(cast(TB_MARKS.YER as integer) between 1 and 3))'));
      // Средняя оценка успеваемости класса
      lCSrClassMark.Caption := FloatToStr(fData.SrClassMark(integer(sgClassNames.Objects[0, sgClassNames.Row])));
     end;
  1: begin
      // Заполняем список учеников
      case fData.FillFIOSg('TB_PEOP','Class',integer(sgClassNames.Objects[0, sgClassNames.Row]), sgClassInfo) of
       0: sgClassInfoClick(self);
       1: sgClassInfoClick(self);
      end;
      // Проверяем кнопки на надобность =) (Есть ли класс)
      if (integer(sgClassNames.Objects[0, sgClassNames.Row]) = -1) then begin bPeopAdd.Visible  := FALSE; bPeopAdd.Enabled  := FALSE; end else begin bPeopAdd.Visible  := TRUE; bPeopAdd.Enabled  := TRUE; end;
      if (fData.Admin = 1) then
      begin
       bPeopAdd.Enabled := FALSE;
       bPeopDel.Enabled := FALSE;
       bPeopMoveClass.Enabled := FALSE;
       bPeopMoveSchool.Enabled := FALSE;
      end;
     end;
  2: rClassInfoMarksPClick(Self);
  3: // Заполняем список учеников
      case fData.FillFIOMarksSG(integer(sgClassNames.Objects[0, sgClassNames.Row]), sgClassMedPeop, 110) of
       0: sgClassMedPeopClick(self);
       1: sgClassMedPeopClick(self);
      end;
  4: // Заполняем список учеников
      case fData.FillFIOSg('TB_PEOP','CLASS',integer(sgClassNames.Objects[0, sgClassNames.Row]), sgAMarksPeop) of
       0: sgAMarksPeopClick(self);
       1: sgAMarksPeopClick(self);
      end;
 end;

end;

procedure TfMain.sgClassInfoClick(Sender: TObject);
var s: string;
begin
 // Проверям кнопки
 if (integer(sgClassInfo.Objects[0, sgClassInfo.Row]) = -1) then begin bPeopDel.Visible  := FALSE; bPeopDel.Enabled  := FALSE; end else begin bPeopDel.Visible  := TRUE; bPeopDel.Enabled  := TRUE; end;
 // Заполняем информацию об ученике и проверяем кнопку
 EPeopFam.Text   := Trim(fData.cSelectS('FNAME','TB_PEOP','Where ID='+IntToStr(integer(sgClassInfo.Objects[0, sgClassInfo.Row]))));
 EPeopName.Text  := Trim(fData.cSelectS('NAME', 'TB_PEOP','Where ID='+IntToStr(integer(sgClassInfo.Objects[0, sgClassInfo.Row]))));
 EPeopSname.Text := Trim(fData.cSelectS('SNAME','TB_PEOP','Where ID='+IntToStr(integer(sgClassInfo.Objects[0, sgClassInfo.Row]))));
 s := fData.cSelectS('BIRTHDAY','TB_PEOP','Where ID='+IntToStr(integer(sgClassInfo.Objects[0, sgClassInfo.Row])));
 if (Trim(s) = '') then ePeopBirthDay.DateTime := StrToDateTime('21.07.1988 1:00:00') else ePeopBirthDay.DateTime := StrToDateTime(s);
 ePeopAdr.Text   := Trim(fData.cSelectS('ADR','TB_PEOP','Where ID='+IntToStr(integer(sgClassInfo.Objects[0, sgClassInfo.Row]))));
 ePeopPS.Text    := Trim(fData.cSelectS('PS','TB_PEOP','Where ID='+IntToStr(integer(sgClassInfo.Objects[0, sgClassInfo.Row]))));
 ePeopPN.Text    := Trim(fData.cSelectS('PN','TB_PEOP','Where ID='+IntToStr(integer(sgClassInfo.Objects[0, sgClassInfo.Row]))));
 ePeopINN.Text   := Trim(fData.cSelectS('INN','TB_PEOP','Where ID='+IntToStr(integer(sgClassInfo.Objects[0, sgClassInfo.Row]))));
 ePeopPrim.Text  := Trim(fData.cSelectS('PRIM','TB_PEOP','Where ID='+IntToStr(integer(sgClassInfo.Objects[0, sgClassInfo.Row]))));
 if (integer(sgClassInfo.Objects[0, sgClassInfo.Row]) = -1) then begin {bSavePeopInfo.Visible  := FALSE;} bSavePeopInfo.Enabled  := FALSE; end else begin {bSavePeopInfo.Visible  := TRUE;} bSavePeopInfo.Enabled  := TRUE; end;
 // Заполняем список классов для перевода (кроме текущего) и проверяем на валидность
 case fData.FillClassCb(cbPeopMoveClass,'where ((SCHOOL='+IntToStr(integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag))+')and(ID<>'+IntToStr(integer(sgClassNames.Objects[0, sgClassNames.Row]))+'))') of
  0: bPeopMoveClass.Enabled := TRUE;
  1: bPeopMoveClass.Enabled := FALSE;
 end;
 // Проверка если нет учеников - вырубаем кнопку перевода по классам
 if (integer(sgClassInfo.Objects[0, sgClassInfo.Row]) = -1) then begin bPeopMoveClass.Enabled := FALSE; bPeopMoveSchool.Enabled := FALSE; end else begin bPeopMoveClass.Enabled := TRUE; bPeopMoveSchool.Enabled := TRUE; end;
 // Заполняем список школ для перевода (кроме текущей) и проверяем на валидность
 case fData.cFillCb(cbPeopMoveSchool,'NAME','TB_SCHOOL','where (ID<>'+IntToStr(integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag))+')') of
  0: cbPeopMoveSchoolChange(self);
  1: cbPeopMoveSchoolChange(self);
 end;
  // Проверка на тип пользователя для кнопки "Добавить"
 if (fData.Admin = 1) then
 begin
  bPeopAdd.Enabled := FALSE;
  bPeopDel.Enabled := FALSE;
  bPeopMoveClass.Enabled := FALSE;
  bPeopMoveSchool.Enabled := FALSE;
 end;
end;

procedure TfMain.PagerChange(Sender: TObject);
var s: string;
begin
 // в зависимости от открытой вкладки - обновляем в ней данные
 case Pager.ActivePageIndex of
  // вкладка школы - статистическая информация
  0: begin
      // Количество учеников
      lCSPeop.Caption := IntToStr(fData.cCount('TB_PEOP.ID','TB_PEOP, TB_CLASS','where ((TB_CLASS.SCHOOL='+IntToStr(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag)+')and(TB_CLASS.NUM between 1 and '+IntToStr(fData.MaxClass-1)+')and(TB_PEOP.CLASS=TB_CLASS.ID))'));
      // Распределение учеников по баллам
      lCSMarks10.Caption := IntToStr(fData.cCount('TB_PEOP.ID','TB_PEOP, TB_MARKS, TB_CLASS','where ((TB_PEOP.ID=TB_MARKS.PEOP)and(TB_CLASS.SCHOOL='+IntToStr(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag)+')and(TB_PEOP.CLASS=TB_CLASS.ID)and(TB_MARKS.FL=0)and(cast(TB_MARKS.YER as integer) between 10 and 12))'));
      lCSMarks7.Caption := IntToStr(fData.cCount('TB_PEOP.ID','TB_PEOP, TB_MARKS, TB_CLASS','where ((TB_PEOP.ID=TB_MARKS.PEOP)and(TB_CLASS.SCHOOL='+IntToStr(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag)+')and(TB_PEOP.CLASS=TB_CLASS.ID)and(TB_MARKS.FL=0)and(cast(TB_MARKS.YER as integer) between 7 and 9))'));
      lCSMarks4.Caption := IntToStr(fData.cCount('TB_PEOP.ID','TB_PEOP, TB_MARKS, TB_CLASS','where ((TB_PEOP.ID=TB_MARKS.PEOP)and(TB_CLASS.SCHOOL='+IntToStr(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag)+')and(TB_PEOP.CLASS=TB_CLASS.ID)and(TB_MARKS.FL=0)and(cast(TB_MARKS.YER as integer) between 4 and 6))'));
      lCSMarks1.Caption := IntToStr(fData.cCount('TB_PEOP.ID','TB_PEOP, TB_MARKS, TB_CLASS','where ((TB_PEOP.ID=TB_MARKS.PEOP)and(TB_CLASS.SCHOOL='+IntToStr(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag)+')and(TB_PEOP.CLASS=TB_CLASS.ID)and(TB_MARKS.FL=0)and(cast(TB_MARKS.YER as integer) between 1 and 3))'));
      // Средняя оценка успеваемости школы
      s := FloatToStr(fData.SrSchoolMark(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag));
      if (Length(s) > 5) then s := copy(s, 1, 5);
      lCSrSMark.Caption := s;
      // Количество учителей
      lCSTeacher.Caption := IntToStr(fData.cCount('ID','TB_TEACHER','where SCHOOL='+IntToStr(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag)));
      // заполняем грид количество и успеваемость по паралелям
      fData.FillSGSchoolStat(sgSchoolStat, tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag);
      // выводим герб школы
      fData.FillImg(Image1, tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag);
      // устанавливаем кнопки активными
      if (tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag = -1) then
       bSchoolEdit.Enabled := FALSE else bSchoolEdit.Enabled := TRUE;
      if (tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag = -1) then
       bSchoolDel.Enabled := FALSE else bSchoolDel.Enabled := TRUE;
      // активность поп-меню школы
      if (tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag = -1) then
       pmSchoolEdit.Enabled := FALSE else pmSchoolEdit.Enabled := TRUE;
      if (tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag = -1) then
       pmSchoolDel.Enabled := FALSE else pmSchoolDel.Enabled := TRUE;
     end;
  // вкладка классов - заполняем список классов
  1: begin
      case fData.FillClassSG(Integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag), sgClassNames, 'and (NUM<'+IntToStr(fData.MaxClass)+')') of
       0: sgClassNamesClick(self);
       1: sgClassNamesClick(self);
      end;
      pClassInfoChange(self);
     end;
  // Вкладка учителя - Заполняем список учителей
  2: begin
      case fData.FillFIOSg('TB_TEACHER','School',integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag), sgTeacher, '', 160) of
       0: sgTeacherClick(self);
       1: sgTeacherClick(self);
      end;
      SetRules(1);
     end;
  // Вкладка уч. планы - Заполняем список классов
  3: begin
      PagerPlans.ActivePageIndex := 0;
      case fData.FillTabsText(tPlansClassName, 'NAME', 'TB_UPLAN', 'where (SCHOOL='+IntToStr(integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag))+')', 'order by NAME') of
       0: tPlansClassNameChange(self);
       1: tPlansClassNameChange(self);
      end;
      SetRules(1);
     end;
  // Вкладка выпускники
  4: case fData.FillClassSG(Integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag), sgEndClasses, 'and (NUM>='+IntToStr(fData.MaxClass)+')') of
      0: sgClassNamesClick(self);
      1: sgClassNamesClick(self);
     end;
 end;
end;

procedure TfMain.sgTeacherClick(Sender: TObject);
var s: string;
begin
  // Если нет школ - вырубаем кнопку Добавления
 if (integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag) = -1) then bTeacherAdd.Enabled  := FALSE else bTeacherAdd.Enabled  := TRUE;
 // Проверям кнопки
 if (integer(sgTeacher.Objects[0, sgTeacher.Row]) = -1) then begin bTeacherDel.Visible  := FALSE; bTeacherDel.Enabled  := FALSE; end else begin bTeacherDel.Visible  := TRUE; bTeacherDel.Enabled  := TRUE; end;
 // Заполняем информацию об ученике и проверяем кнопку
 eTeacherFname.Text := Trim(fData.cSelectS('FNAME','TB_TEACHER','Where ID='+IntToStr(integer(sgTeacher.Objects[0, sgTeacher.Row]))));
 eTeacherName.Text  := Trim(fData.cSelectS('NAME', 'TB_TEACHER','Where ID='+IntToStr(integer(sgTeacher.Objects[0, sgTeacher.Row]))));
 eTeacherSname.Text := Trim(fData.cSelectS('SNAME','TB_TEACHER','Where ID='+IntToStr(integer(sgTeacher.Objects[0, sgTeacher.Row]))));
 s := fData.cSelectS('BIRTHDAY','TB_TEACHER','Where ID='+IntToStr(integer(sgTeacher.Objects[0, sgTeacher.Row])));
 if (Trim(s) = '') then eTeacherBirthDay.DateTime := StrToDateTime('21.07.1988 1:00:00') else eTeacherBirthDay.DateTime := StrToDateTime(s);
 eTeacherAdr.Text := Trim(fData.cSelectS('ADR','TB_TEACHER','Where ID='+IntToStr(integer(sgTeacher.Objects[0, sgTeacher.Row]))));
 eTeacherPS.Text := Trim(fData.cSelectS('PS','TB_TEACHER','Where ID='+IntToStr(integer(sgTeacher.Objects[0, sgTeacher.Row]))));
 eTeacherPN.Text := Trim(fData.cSelectS('PN','TB_TEACHER','Where ID='+IntToStr(integer(sgTeacher.Objects[0, sgTeacher.Row]))));
 eTeacherINN.Text := Trim(fData.cSelectS('INN','TB_TEACHER','Where ID='+IntToStr(integer(sgTeacher.Objects[0, sgTeacher.Row]))));
 eTeacherPrim.Text := Trim(fData.cSelectS('PRIM','TB_TEACHER','Where ID='+IntToStr(integer(sgTeacher.Objects[0, sgTeacher.Row]))));

 if (integer(sgTeacher.Objects[0, sgTeacher.Row]) = -1) then begin {bSavePeopInfo.Visible  := FALSE;} bTeacherSave.Enabled  := FALSE; end else begin {bSavePeopInfo.Visible  := TRUE;} bTeacherSave.Enabled  := TRUE; end;
 // Заполняем список школ для перевода (кроме текущей) и проверяем на валидность
 case fData.cFillCb(cbTeacherMoveSchool,'NAME','TB_SCHOOL','where (ID<>'+IntToStr(integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag))+')') of
  0: if (integer(sgTeacher.Objects[0, sgTeacher.Row]) = -1) then begin {bTeacherDel.Visible  := FALSE;} bTeacherMoveSchool.Enabled  := FALSE; end else begin {bTeacherDel.Visible  := TRUE;} bTeacherMoveSchool.Enabled  := TRUE; end;
  1: // Ненадо выключать кнопку т.к. школы может не быть в списке (вариант программы для 1ой конкретной школы)
     // Выключаем только в случаем, если и переводить то некого ))
     if (integer(sgTeacher.Objects[0, sgTeacher.Row]) = -1) then begin {bTeacherDel.Visible  := FALSE;} bTeacherMoveSchool.Enabled  := FALSE; end else begin {bTeacherDel.Visible  := TRUE;} bTeacherMoveSchool.Enabled  := TRUE; end;
 end;
 SetRules(1);
end;

procedure TfMain.bPeopAddClick(Sender: TObject);
var r,v: TStringList;
begin
  r := TStringList.Create;
  v := TStringList.Create;
 try
  r.Clear; v.Clear;
  r.Add('Class'); r.Add('FName'); r.Add('Name'); r.Add('SName'); //r.Add('LOG');
  v.Add(IntToStr(integer(sgClassNames.Objects[0, sgClassNames.Row])));
  v.Add(''''+'*'+fData.standarts[2]+'''');
  v.Add(''''+fData.standarts[2]+'''');
  v.Add(''''+fData.standarts[2]+'''');
 // v.Add(''''+DateToStr(Now())+' '+fData.standarts[3]+' '+fData.standarts[4]+' '+tClassNames.AdvOfficeTabs[tClassNames.ActiveTabIndex].Caption+'''');
  // Подготовили новую запись и отправили на выполнение
  case fData.cInserts('TB_PEOP',r,v) of
      // Успешно добавилось
   0: pClassInfoChange(Self);
   1:
  end;
 finally
  r.Free; v.Free;
 end;
end;

procedure TfMain.bSavePeopInfoClick(Sender: TObject);
var r,v: TStringList;
begin
 r := TStringList.Create;
 v := TStringList.Create;
 try
  r.Clear; v.Clear;
  r.Add('FName'); r.Add('Name'); r.Add('SName'); r.Add('BIRTHDAY');
  r.Add('ADR'); r.Add('PS'); r.Add('PN'); r.Add('INN'); r.Add('PRIM');
  v.Add(''''+Trim(EPeopFam.Text)+'''');
  v.Add(''''+Trim(EPeopName.Text)+'''');
  v.Add(''''+Trim(EPeopSname.Text)+'''');
  v.Add(''''+DateToStr(ePeopBirthDay.Date)+'''');
  v.Add(''''+Trim(ePeopAdr.Text)+'''');
  v.Add(''''+Trim(ePeopPS.Text)+'''');
  if (Trim(ePeopPN.Text) = '') then v.Add('0') else
  v.Add(Trim(ePeopPN.Text));
  v.Add(''''+Trim(ePeopINN.Text)+'''');
  v.Add(''''+Trim(ePeopPrim.Text)+'''');
  // Подготовили новую запись и отправили на выполнение
  case fData.cUpdates('TB_PEOP',r,v,'where ID='+IntToStr(integer(sgClassInfo.Objects[0, sgClassInfo.Row]))) of
      // Успешно добавилось
   0: pClassInfoChange(Self);
   1:
  end;
 finally
  r.Free; v.Free;
 end;
end;

procedure TfMain.bTeacherAddClick(Sender: TObject);
var r,v: TStringList;
begin
 r := TStringList.Create;
 v := TStringList.Create;
 try
  r.Clear; v.Clear;
  r.Add('School'); r.Add('FName'); r.Add('Name'); r.Add('SName');
  v.Add(IntToStr(integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag)));
  v.Add(''''+'*'+fData.standarts[2]+'''');
  v.Add(''''+fData.standarts[2]+'''');
  v.Add(''''+fData.standarts[2]+'''');
//  v.Add(''''+DateToStr(Now())+' '+fData.standarts[3]+' '+fData.standarts[4]+' '+sgSchools.Cells[0, sgSchools.Row]+'''');
  // Подготовили новую запись и отправили на выполнение
  case fData.cInserts('TB_TEACHER',r,v) of
      // Успешно добавилось
   0: PagerChange(Self);
   1:
  end;
 finally
  r.Free; v.Free;
 end;

end;

procedure TfMain.bTeacherSaveClick(Sender: TObject);
var r,v: TStringList;
begin
 r := TStringList.Create;
 v := TStringList.Create;
 try
  r.Clear; v.Clear;
  r.Add('FName'); r.Add('Name'); r.Add('SName'); r.Add('BIRTHDAY');
  r.Add('ADR'); r.Add('PS'); r.Add('PN'); r.Add('INN'); r.Add('PRIM');
  v.Add(''''+Trim(eTeacherFname.Text)+'''');
  v.Add(''''+Trim(eTeacherName.Text)+'''');
  v.Add(''''+Trim(eTeacherSname.Text)+'''');
  v.Add(''''+DateToStr(eTeacherBirthDay.Date)+'''');
  v.Add(''''+Trim(eTeacherAdr.Text)+'''');
  v.Add(''''+Trim(eTeacherPS.Text)+'''');
  v.Add(Trim(eTeacherPN.Text));
  v.Add(''''+Trim(eTeacherINN.Text)+'''');
  v.Add(''''+Trim(eTeacherPrim.Text)+'''');
  // Подготовили новую запись и отправили на выполнение
  case fData.cUpdates('TB_TEACHER',r,v,'where ID='+IntToStr(integer(sgTeacher.Objects[0, sgTeacher.Row]))) of
      // Успешно добавилось
   0: PagerChange(Self);
   1:
  end;
 finally
  r.Free; v.Free;
 end;
end;

procedure TfMain.tPlansClassNameChange(Sender: TObject);
begin
 if ((tPlansClassName.AdvOfficeTabs.Count=1)and(tPlansClassName.AdvOfficeTabs[tPlansClassName.ActiveTabIndex].Tag=-1)) then bPlanDel.Enabled := FALSE else bPlanDel.Enabled := TRUE;
 if ((tPlansClassName.AdvOfficeTabs.Count=1)and(tPlansClassName.AdvOfficeTabs[tPlansClassName.ActiveTabIndex].Tag=-1)) then bPlanEdit.Enabled := FALSE else bPlanEdit.Enabled := TRUE;
 UPlanFilter := 0;
 UPlanFilterID := tPlansClassName.AdvOfficeTabs[tPlansClassName.ActiveTabIndex].Tag;
 bPlanAdd.Enabled := TRUE;
 // Заполняем список учителей
 case fData.FillFIOCb('TB_TEACHER','School',integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag), cbPlansTeach) of
  1: bPlanAdd.Enabled := FALSE;
 end;
 // Заполняем список предметов, кот. нет в уч. плане текущего класса
 case fData.FillCbPredmetPlan(cbPlansPr,tPlansClassName.AdvOfficeTabs[tPlansClassName.ActiveTabIndex].Tag) of
  1: bPlanAdd.Enabled := FALSE;
 end;
 // Заполняем уч. план для текущего класса
 case fData.FillPlanSg(tPlansClassName.AdvOfficeTabs[tPlansClassName.ActiveTabIndex].Tag,sgPlansClass) of
  0:
 end;
 // Проверяем кнопки
 if (tPlansClassName.AdvOfficeTabs[tPlansClassName.ActiveTabIndex].Tag=-1) then
 begin
  bPlanRAdd.Enabled := FALSE;
  bPlanRDel.Enabled := FALSE;
  bPlanImport.Enabled := FALSE;
 end else
 begin
  bPlanRAdd.Enabled := TRUE;
  bPlanRDel.Enabled := TRUE;
  bPlanImport.Enabled := TRUE;
 end;
end;

procedure TfMain.bPlanRAddClick(Sender: TObject);
var r,v: TStringList;
begin
 r := TStringList.Create;
 v := TStringList.Create;
 try
  r.Clear; v.Clear;
  r.Add('UPLAN'); r.Add('Teacher'); r.Add('Predmet');
  v.Add(IntToStr(tPlansClassName.AdvOfficeTabs[tPlansClassName.ActiveTabIndex].Tag));
  v.Add(IntToStr(integer(cbPlansTeach.Items.Objects[cbPlansTeach.ItemIndex])));
  v.Add(IntToStr(integer(cbPlansPr.Items.Objects[cbPlansPr.ItemIndex])));
  // Подготовили новую запись и отправили на выполнение
  case fData.cInserts('TB_PLANCLASS',r,v) of
      // Успешно добавилось
   0: tPlansClassNameChange(Self);
   1:
  end;
 finally
  r.Free; v.Free;
 end;
end;

procedure TfMain.bPredmetClick(Sender: TObject);
begin
 Application.CreateForm(TfPredmet, fPredmet);
 fPredmet.Show();
end;

procedure TfMain.bPlanRDelClick(Sender: TObject);
begin
 // Проверяем уверенность юзера в своих намерениях =)
 if (MessageDlg(fData.standarts[9],mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
 case fData.cDelete('TB_PLANCLASS','where ID='+IntToStr(integer(sgPlansClass.Objects[0, sgPlansClass.Row]))) of
  0: tPlansClassNameChange(self);
 end;
end;

procedure TfMain.bTeacherDelClick(Sender: TObject);
begin
 // Проверям ведет ли что-то этот учитель
 case fData.ifCount('TEACHER','TB_PLANCLASS','WHERE TEACHER='+IntToStr(integer(sgTeacher.Objects[0, sgTeacher.Row]))) of
  1: fData.ShowError(8,fMain.Handle);
  0: // Проверяем уверенность юзера в своих намерениях =)
     if (MessageDlg(fData.standarts[9],mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
     case fData.cDelete('TB_TEACHER','where ID='+IntToStr(integer(sgTeacher.Objects[0, sgTeacher.Row]))) of
      0: PagerChange(self);
     end;
 end;
end;

procedure TfMain.bPeopDelClick(Sender: TObject);
begin
 // Проверяем уверенность юзера в своих намерениях =)
 if (MessageDlg(fData.standarts[9],mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
 case fData.PeopDel(integer(sgClassInfo.Objects[0, sgClassInfo.Row])) of
  0: sgClassNamesClick(self);
  1: sgClassNamesClick(self);
 end;
end;

procedure TfMain.bClassAddClick(Sender: TObject);
begin
 Application.CreateForm(TfNewClass, fNewClass);
 fNewClass.Mode := 0;
 fNewClass.School := integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag);
 fNewClass.Show();
end;

procedure TfMain.bClassEditClick(Sender: TObject);
begin
 Application.CreateForm(TfNewClass, fNewClass);
 fNewClass.Mode := 1;
 fNewClass.ID := integer(sgClassNames.Objects[0, sgClassNames.Row]);
 fNewClass.School := integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag);
 fNewClass.Show();
end;

procedure TfMain.PagerPlansChange(Sender: TObject);
begin
 case PagerPlans.ActivePageIndex of
  0: tPlansClassNameChange(self);
  1: begin
      bPPlansFilterPr.Enabled := TRUE;
      bPPlansFilterTch.Enabled := TRUE;
      // Заполняем список учителей
      case fData.FillFIOCb('TB_TEACHER','School',integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag), ePPlansFilterTch) of
       1: bPPlansFilterTch.Enabled := FALSE;
      end;
      // Заполняем список предметов
      case fData.cFillCb(ePPlansFilterPr,'NAME','TB_PREDMET') of
       1: bPPlansFilterPr.Enabled := FALSE;
      end;
     end;
 end;
end;

procedure TfMain.bPPlansFilterPrClick(Sender: TObject);
begin
 UPlanFilter := 1;
 UPlanFilterID := integer(ePPlansFilterPr.Items.Objects[ePPlansFilterPr.ItemIndex]);
 case fData.FillPlanSGf(0,integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag),integer(ePPlansFilterPr.Items.Objects[ePPlansFilterPr.ItemIndex]),sgPlansClass) of
  0:
 end;
end;

procedure TfMain.bPPlansFilterTchClick(Sender: TObject);
begin
 UPlanFilter := 2;
 UPlanFilterID := integer(ePPlansFilterTch.Items.Objects[ePPlansFilterTch.ItemIndex]);
 case fData.FillPlanSGf(1,integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag),integer(ePPlansFilterTch.Items.Objects[ePPlansFilterTch.ItemIndex]),sgPlansClass) of
  0:
 end;
end;

procedure TfMain.bPPlansReportExcelClick(Sender: TObject);
begin
 case UPlanFilter of
  0: fData.ExpPlansExcel(0, UPlanFilterID, sgPlansClass);
  1: fData.ExpPlansExcel(1, UPlanFilterID, sgPlansClass);
  2: fData.ExpPlansExcel(2, UPlanFilterID, sgPlansClass);
 end;
end;

procedure TfMain.bPlanImportClick(Sender: TObject);
begin
 Application.CreateForm(TfPlanImport, fPlanImport);
 fPlanImport.School := integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag);
 fPlanImport.ID := tPlansClassName.AdvOfficeTabs[tPlansClassName.ActiveTabIndex].Tag;
 fPlanImport.Show();
end;

procedure TfMain.bClassDelClick(Sender: TObject);
begin
 // Проверяем уверенность юзера в своих намерениях =)
 if (MessageDlg(fData.standarts[9],mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
 begin
  fData.ClassDel(integer(sgClassNames.Objects[0, sgClassNames.Row]));
  PagerChange(Self);
 end;
end;

procedure TfMain.rClassInfoMarksPClick(Sender: TObject);
begin
 if (rClassInfoMarksP.Checked) then fData.FillFIOMarksSG(integer(sgClassNames.Objects[0, sgClassNames.Row]), sgClassInfoMarksAll) else fData.FillPrMarksSG(integer(sgClassNames.Objects[0, sgClassNames.Row]), sgClassInfoMarksAll);
 sgClassInfoMarksAllClick(Self);
end;

procedure TfMain.rClassInfoMarksWClick(Sender: TObject);
begin
  if (rClassInfoMarksP.Checked) then fData.FillFIOMarksSG(integer(sgClassNames.Objects[0, sgClassNames.Row]), sgClassInfoMarksAll) else fData.FillPrMarksSG(integer(sgClassNames.Objects[0, sgClassNames.Row]), sgClassInfoMarksAll);
end;

procedure TfMain.sgClassInfoMarksAllClick(Sender: TObject);
begin
 if (integer(sgClassInfoMarksAll.Objects[0, sgClassInfoMarksAll.Row]) = -1) then
  case fData.FillMarksAll(integer(sgClassNames.Objects[0, sgClassNames.Row]), sgClassInfoMarks) of
   0: begin bMarksSave.Enabled := TRUE;  bMarksExcel.Enabled := TRUE;  bMarksChart.Enabled := TRUE;  end;
   1: begin bMarksSave.Enabled := FALSE; bMarksExcel.Enabled := FALSE; bMarksChart.Enabled := FALSE; end;
  end else
 if (rClassInfoMarksP.Checked) then
  case fData.FillMarksPp(integer(sgClassNames.Objects[0, sgClassNames.Row]), integer(sgClassInfoMarksAll.Objects[0, sgClassInfoMarksAll.Row]), sgClassInfoMarks) of
   0: begin bMarksSave.Enabled := TRUE;  bMarksExcel.Enabled := TRUE;  bMarksChart.Enabled := TRUE;  end;
   1: begin bMarksSave.Enabled := FALSE; bMarksExcel.Enabled := FALSE; bMarksChart.Enabled := FALSE; end;
  end else
  case fData.FillMarksPr(integer(sgClassNames.Objects[0, sgClassNames.Row]), integer(sgClassInfoMarksAll.Objects[0, sgClassInfoMarksAll.Row]), sgClassInfoMarks) of
   0: begin bMarksSave.Enabled := TRUE;  bMarksExcel.Enabled := TRUE;  bMarksChart.Enabled := TRUE;  end;
   1: begin bMarksSave.Enabled := FALSE; bMarksExcel.Enabled := FALSE; bMarksChart.Enabled := FALSE; end;
  end;
  // Проверка на тип пользователя для кнопки "Добавить"
 if (fData.Admin = 1) then bMarksSave.Enabled := FALSE;
end;

procedure TfMain.sgClassInfoMarksClick(Sender: TObject);
begin
// ShowMessage('col= '+IntToStr(sgClassInfoMarks.Col)+'   row= '+IntToStr(sgClassInfoMarks.Row));
end;

procedure TfMain.bMarksSaveClick(Sender: TObject);
begin
 if (sgClassInfoMarksAll.Row = 1) then fData.MarksAllSave(sgClassInfoMarks) else
  if (rClassInfoMarksP.Checked) then fData.MarksAllSaveP(0, sgClassInfoMarks) else fData.MarksAllSaveP(1, sgClassInfoMarks);
 sgClassInfoMarksAllClick(Self);
end;

procedure TfMain.bMarksExcelClick(Sender: TObject);
begin
 if (sgClassInfoMarksAll.Row = 1) then fData.ExpMarksExcelAll(sgClassInfoMarks) else
  if (rClassInfoMarksP.Checked) then fData.ExpMarksExcelP(0, sgClassInfoMarks) else fData.ExpMarksExcelP(1, sgClassInfoMarks);
end;

procedure TfMain.bMarksChartClick(Sender: TObject);
begin
 if (sgClassInfoMarksAll.Row = 1) then fData.ExpMarksDiag(0, sgClassInfoMarks) else
  if (rClassInfoMarksP.Checked) then fData.ExpMarksDiag(1, sgClassInfoMarks) else fData.ExpMarksDiag(2, sgClassInfoMarks);
end;

procedure TfMain.MnClassInClick(Sender: TObject);
begin
 Application.CreateForm(TfClassOb, fClassOb);
 fClassOb.School := integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag);
 fClassOb.Show();
end;

procedure TfMain.MnClassOutClick(Sender: TObject);
begin
 Application.CreateForm(TfClassRz, fClassRz);
 fClassRz.School := integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag);
 fClassRz.Show();
end;

procedure TfMain.bPlanAddClick(Sender: TObject);
begin
 Application.CreateForm(TfUPlanAdd, fUPlanAdd);
 fUPlanAdd.Mode := 0;
 fUPlanAdd.School := integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag);
 fUPlanAdd.Show();
end;

procedure TfMain.bPlanEditClick(Sender: TObject);
begin
 Application.CreateForm(TfUPlanAdd, fUPlanAdd);
 fUPlanAdd.Mode := 1;
 fUPlanAdd.School := integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag);
 fUPlanAdd.ID := tPlansClassName.AdvOfficeTabs[tPlansClassName.ActiveTabIndex].Tag;
 fUPlanAdd.Show();
end;

procedure TfMain.sgClassMedPeopClick(Sender: TObject);
var filter: string;
begin
 if (integer(sgClassMedPeop.Objects[0, sgClassMedPeop.Row]) = -1) then
  bClassMedAdd.Enabled := FALSE else bClassMedAdd.Enabled := TRUE;
 if (cClassMedFilter.Checked) then filter := 'and(TB_MEDICINA.DB >= '''+DateToStr(eClassMedFilter.Date)+''')' else filter := '';
 if (sgClassMedPeop.Row = 1) then
  fData.FillMedAll(integer(sgClassNames.Objects[0, sgClassNames.Row]), sgClassMed, filter)
 else fData.FillMed(integer(sgClassMedPeop.Objects[0, sgClassMedPeop.Row]), sgClassMed, filter);
 // Проверка на тип пользователя для кнопки "Добавить"
 if (fData.Admin = 1) then bClassMedAdd.Enabled := FALSE;
end;

procedure TfMain.tClassInfoMedShow(Sender: TObject);
begin
 eClassMedFilter.DateTime := Now();
 eClassMedDB.DateTime := Now();
 eClassMedDE.DateTime := Now();
 eClassMedTxt.Text := '';
end;

procedure TfMain.cClassMedFilterClick(Sender: TObject);
begin
 sgEndMedPeopClick(self);
end;

procedure TfMain.bClassMedAddClick(Sender: TObject);
var r,v: TStringList;
begin
 if (Length(Trim(eClassMedTxt.Text)) = 0) then
 begin
  fData.ShowError(11);
  exit;
 end;
 if (eClassMedDB.Date = eClassMedDE.Date)or(eClassMedDB.Date > eClassMedDE.Date) then
 begin
  fData.ShowError(12);
  exit;
 end;
 r := TStringList.Create(); r.Clear();
 v := TStringList.Create(); v.Clear();
 r.Add('ID'); r.Add('PEOP'); r.Add('DB'); r.Add('DE'); r.Add('TXT');
 v.Add(IntToStr(fData.cMax('ID','TB_MEDICINA')+1));
 v.Add(IntToStr(integer(sgClassMedPeop.Objects[0, sgClassMedPeop.Row])));
 v.Add(''''+DateToStr(eClassMedDB.Date)+'''');
 v.Add(''''+DateToStr(eClassMedDE.Date)+'''');
 v.Add(''''+Trim(eClassMedTxt.Text)+'''');
 if (fData.cInserts('TB_MEDICINA', r,v) = 0) then
 begin
  tClassInfoMedShow(self);
  sgClassMedPeopClick(self);
 end;
 r.Free; v.Free;
end;

procedure TfMain.sgClassNamesClick(Sender: TObject);
begin
 // Если нет школ - вырубаем кнопку Добавления
 if (integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag) = -1) then bClassAdd.Enabled  := FALSE else bClassAdd.Enabled  := TRUE;
 // Если классов нет - вырубаем кнопки
 if (integer(sgClassNames.Objects[0, sgClassNames.Row]) = -1) then bClassEdit.Enabled  := FALSE else bClassEdit.Enabled  := TRUE;
 if (integer(sgClassNames.Objects[0, sgClassNames.Row]) = -1) then bClassDel.Enabled   := FALSE else bClassDel.Enabled   := TRUE;
 if (integer(sgClassNames.Objects[0, sgClassNames.Row]) = -1) then bClassIn.Enabled    := FALSE else bClassIn.Enabled    := TRUE;
 if (integer(sgClassNames.Objects[0, sgClassNames.Row]) = -1) then bClassBreak.Enabled := FALSE else bClassBreak.Enabled := TRUE;
 if (integer(sgClassNames.Objects[0, sgClassNames.Row]) = -1) then pmClassEdit.Enabled  := FALSE else pmClassEdit.Enabled  := TRUE;
 if (integer(sgClassNames.Objects[0, sgClassNames.Row]) = -1) then pmClassDel.Enabled   := FALSE else pmClassDel.Enabled   := TRUE;
 if (integer(sgClassNames.Objects[0, sgClassNames.Row]) = -1) then pmClassIn.Enabled    := FALSE else pmClassIn.Enabled    := TRUE;
 if (integer(sgClassNames.Objects[0, sgClassNames.Row]) = -1) then pmClassBreak.Enabled := FALSE else pmClassBreak.Enabled := TRUE;
 if (integer(sgClassNames.Objects[0, sgClassNames.Row]) = -1) then bClassPlan.Enabled := FALSE else bClassPlan.Enabled := TRUE;

 // При изменении текущего класса изменяем данные вкладки
 pClassInfoChange(self);
 if (fData.Admin = 1) then
 begin
  bPeopAdd.Enabled := FALSE;
  bPeopDel.Enabled := FALSE;
  bPeopMoveClass.Enabled := FALSE;
  bPeopMoveSchool.Enabled := FALSE;
 end;
end;

procedure TfMain.tSchoolsTabClick(Sender: TObject; PageIndex: Integer);
begin
 PagerChange(self);
end;

procedure TfMain.bClassPlanClick(Sender: TObject);
begin
 fData.ChangePlan(integer(sgClassNames.Objects[0, sgClassNames.Row]), integer(cbClassPlan.Items.Objects[cbClassPlan.ItemIndex]));
end;

procedure TfMain.sgAMarksPeopClick(Sender: TObject);
begin
 case fData.FillAClassCb(cClassAMarksClass, integer(sgAMarksPeop.Objects[0, sgAMarksPeop.Row])) of
  0: begin
      bClassAMarksClass.Enabled := TRUE;
      pAMarksExcel.Enabled := TRUE;
      pAMarksChart.Enabled := TRUE;
      if (integer(cClassAMarksClass.Items.Objects[cClassAMarksClass.ItemIndex]) <> -1) then bClassAMarksClassClick(self) else
      begin
       sgAMarksM.RowCount := 1;
       sgAMarksM.ColCount := 1;
      end;
     end;
  1: begin;
      bClassAMarksClass.Enabled := FALSE;
      pAMarksExcel.Enabled := FALSE;
      pAMarksChart.Enabled := FALSE;
      sgAMarksM.RowCount := 1;
      sgAMarksM.ColCount := 1;
     end;
 end;
end;

procedure TfMain.bClassAMarksClassClick(Sender: TObject);
begin
 fData.FillAMarks(integer(sgAMarksPeop.Objects[0, sgAMarksPeop.Row]), StrToInt(cClassAMarksClass.Items[cClassAMarksClass.ItemIndex]), sgAMarksM);
end;

procedure TfMain.pAMarksExcelClick(Sender: TObject);
begin
 fData.ExpMarksExcelP(0, sgAMarksM);
end;

procedure TfMain.pAMarksChartClick(Sender: TObject);
begin
 fData.ExpMarksDiag(1, sgAMarksM);
end;

procedure TfMain.sgEndClassesClick(Sender: TObject);
begin
 pEndNavChange(self); 
end;

procedure TfMain.pEndNavChange(Sender: TObject);
begin
 // В зависимости от открытой вкладки - обновляем инфу
 case pEndNav.ActivePageIndex of
  0: begin
      // Количество учеников
      lEndInfoCountR.Caption := IntToStr(fData.cCount('ID','TB_PEOP','where CLASS='+IntToStr(integer(sgEndClasses.Objects[0, sgEndClasses.Row]))));
            // Количество учеников
      // Распределение учеников по баллам
      Label9.Caption := IntToStr(fData.cCount('TB_PEOP.ID','TB_PEOP, TB_AMARKS','where ((TB_PEOP.ID=TB_AMARKS.PEOP)and(TB_PEOP.CLASS='+IntToStr(integer(sgEndClasses.Objects[0, sgEndClasses.Row]))+')and(TB_AMARKS.FL=0)and(cast(TB_AMARKS.YER as integer) between 10 and 12))'));
      Label8.Caption := IntToStr(fData.cCount('TB_PEOP.ID','TB_PEOP, TB_AMARKS','where ((TB_PEOP.ID=TB_AMARKS.PEOP)and(TB_PEOP.CLASS='+IntToStr(integer(sgEndClasses.Objects[0, sgEndClasses.Row]))+')and(TB_AMARKS.FL=0)and(cast(TB_AMARKS.YER as integer) between 7 and 9))'));
      Label5.Caption := IntToStr(fData.cCount('TB_PEOP.ID','TB_PEOP, TB_AMARKS','where ((TB_PEOP.ID=TB_AMARKS.PEOP)and(TB_PEOP.CLASS='+IntToStr(integer(sgEndClasses.Objects[0, sgEndClasses.Row]))+')and(TB_AMARKS.FL=0)and(cast(TB_AMARKS.YER as integer) between 4 and 6))'));
      Label4.Caption := IntToStr(fData.cCount('TB_PEOP.ID','TB_PEOP, TB_AMARKS','where ((TB_PEOP.ID=TB_AMARKS.PEOP)and(TB_PEOP.CLASS='+IntToStr(integer(sgEndClasses.Objects[0, sgEndClasses.Row]))+')and(TB_AMARKS.FL=0)and(cast(TB_AMARKS.YER as integer) between 1 and 3))'));
      // Средняя оценка успеваемости класса
      Label2.Caption := FloatToStr(fData.SrClassAMark(integer(sgEndClasses.Objects[0, sgEndClasses.Row])));

     end;
  1: begin
      // Заполняем список учеников
      case fData.FillFIOSg('TB_PEOP','CLASS',integer(sgEndClasses.Objects[0, sgEndClasses.Row]), sgEndInfoPeop) of
       0: sgEndInfoPeopClick(self);
       1: sgEndInfoPeopClick(self);
      end;
     end;
  2: // Заполняем список учеников
      case fData.FillFIOSg('TB_PEOP','CLASS',integer(sgEndClasses.Objects[0, sgEndClasses.Row]), sgEndAMarksPeop) of
       0: sgEndAMarksPeopClick(self);
       1: sgEndAMarksPeopClick(self);
      end;
  3: // Заполняем список учеников
      case fData.FillFIOMarksSG(integer(sgEndClasses.Objects[0, sgEndClasses.Row]), sgEndMedPeop, 110) of
       0: sgEndMedPeopClick(self);
       1: sgEndMedPeopClick(self);
      end;
 end;
end;

procedure TfMain.sgEndInfoPeopClick(Sender: TObject);
var s: string;
begin
 // Заполняем информацию об ученике и проверяем кнопку
 epEndInfoFName.Text := Trim(fData.cSelectS('FNAME','TB_PEOP','Where ID='+IntToStr(integer(sgEndInfoPeop.Objects[0, sgEndInfoPeop.Row]))));
 epEndInfoName.Text  := Trim(fData.cSelectS('NAME', 'TB_PEOP','Where ID='+IntToStr(integer(sgEndInfoPeop.Objects[0, sgEndInfoPeop.Row]))));
 epEndInfoSName.Text := Trim(fData.cSelectS('SNAME','TB_PEOP','Where ID='+IntToStr(integer(sgEndInfoPeop.Objects[0, sgEndInfoPeop.Row]))));

 s := fData.cSelectS('BIRTHDAY','TB_PEOP','Where ID='+IntToStr(integer(sgEndInfoPeop.Objects[0, sgEndInfoPeop.Row])));
 if (Trim(s) = '') then DateTimePicker1.DateTime := StrToDateTime('21.07.1988 1:00:00') else DateTimePicker1.DateTime := StrToDateTime(s);
 AdvEdit1.Text   := Trim(fData.cSelectS('ADR','TB_PEOP','Where ID='+IntToStr(integer(sgEndInfoPeop.Objects[0, sgEndInfoPeop.Row]))));
 AdvEdit2.Text    := Trim(fData.cSelectS('PS','TB_PEOP','Where ID='+IntToStr(integer(sgEndInfoPeop.Objects[0, sgEndInfoPeop.Row]))));
 AdvEdit3.Text    := Trim(fData.cSelectS('PN','TB_PEOP','Where ID='+IntToStr(integer(sgEndInfoPeop.Objects[0, sgEndInfoPeop.Row]))));
 AdvEdit4.Text   := Trim(fData.cSelectS('INN','TB_PEOP','Where ID='+IntToStr(integer(sgEndInfoPeop.Objects[0, sgEndInfoPeop.Row]))));
 Memo1.Text  := Trim(fData.cSelectS('PRIM','TB_PEOP','Where ID='+IntToStr(integer(sgEndInfoPeop.Objects[0, sgEndInfoPeop.Row]))));

end;

procedure TfMain.sgEndAMarksPeopClick(Sender: TObject);
begin
 case fData.FillAClassCb(eEndAMarksClass, integer(sgEndAMarksPeop.Objects[0, sgEndAMarksPeop.Row])) of
  0: bEndAMarksClass.Enabled := TRUE;
  1: bEndAMarksClass.Enabled := FALSE;
 end;
end;

procedure TfMain.sgEndMedPeopClick(Sender: TObject);
var filter: string;
begin
 if (cEndMedFilter.Checked) then filter := 'and(TB_MEDICINA.DB >= '''+DateToStr(eEndMedFilter.Date)+''')' else filter := '';
 if (sgEndMedPeop.Row = 1) then
  fData.FillMedAll(integer(sgEndClasses.Objects[0, sgEndClasses.Row]), sgEndMed, filter)
 else fData.FillMed(integer(sgEndMedPeop.Objects[0, sgEndMedPeop.Row]), sgEndMed, filter);
end;

procedure TfMain.bEndAMarksClassClick(Sender: TObject);
begin
 fData.FillAMarks(integer(sgEndAMarksPeop.Objects[0, sgEndAMarksPeop.Row]), StrToInt(eEndAMarksClass.Items[eEndAMarksClass.ItemIndex]), sgEndAMarksM);
end;

procedure TfMain.bEndAMarksExcelClick(Sender: TObject);
begin
 fData.ExpMarksExcelP(0, sgEndAMarksM);
end;

procedure TfMain.bEndAMarksChartClick(Sender: TObject);
begin
 fData.ExpMarksDiag(1, sgEndAMarksM);
end;

procedure TfMain.sgClassInfoMarksKeyPress(Sender: TObject; var Key: Char);
var t: string;
    i: integer;
begin
 if ((Key <> '1')and(Key <> '2')and(Key <> '3')and(Key <> '4')and(Key <> '5')and(Key <> '6')and(Key <> '7')and(Key <> '8')and(Key <> '9')and(Key <> '0')and(Key <> '.')) then
 begin
  t := sgClassInfoMarks.Cells[sgClassInfoMarks.Col, sgClassInfoMarks.Row];
  i := pos(Key, t);
  if (i <> 0) then
  begin
   delete(t, i-1, 1);
   t := copy(t, 1, i-1) + '.' + copy(t, i+1, Length(t));
  end;
  sgClassInfoMarks.Cells[sgClassInfoMarks.Col, sgClassInfoMarks.Row] := t;
 end;
end;

procedure TfMain.MnSchoolAddClick(Sender: TObject);
begin
 Application.CreateForm(TfSchool, fSchool);
 fSchool.Mode := 0;
 fSchool.Show();
end;

procedure TfMain.MnSchoolEditClick(Sender: TObject);
begin
 Application.CreateForm(TfSchool, fSchool);
 fSchool.Mode := 1;
 fSchool.ID := tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag;
 fSchool.Show();
end;

procedure TfMain.MnSchoolDelClick(Sender: TObject);
begin
 fData.DeleteSchool(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag);
 FormShow(self);
end;

procedure TfMain.MnClassAddClick(Sender: TObject);
begin
 Application.CreateForm(TfNewClass, fNewClass);
 fNewClass.Mode := 0;
 fNewClass.School := integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag);
 fNewClass.Show(); 
end;

procedure TfMain.MnClassEditClick(Sender: TObject);
begin
 Application.CreateForm(TfNewClass, fNewClass);
 fNewClass.Mode := 1;
 fNewClass.ID := integer(sgClassNames.Objects[0, sgClassNames.Row]);
 fNewClass.School := integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag);
 fNewClass.Show();
end;

procedure TfMain.MnClassDelClick(Sender: TObject);
begin
 // Проверяем уверенность юзера в своих намерениях =)
 if (MessageDlg(fData.standarts[9],mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
 begin
  fData.ClassDel(integer(sgClassNames.Objects[0, sgClassNames.Row]));
  PagerChange(Self);
 end;
end;

procedure TfMain.bClassInClick(Sender: TObject);
begin
 Application.CreateForm(TfClassOb, fClassOb);
 fClassOb.School := integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag);
 fClassOb.Show();
end;

procedure TfMain.bClassBreakClick(Sender: TObject);
begin
 Application.CreateForm(TfClassRz, fClassRz);
 fClassRz.School := integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag);
 fClassRz.Show();
end;

procedure TfMain.bSchoolAddClick(Sender: TObject);
begin
 Application.CreateForm(TfSchool, fSchool);
 fSchool.Mode := 0;
 fSchool.Show();
end;

procedure TfMain.bSchoolEditClick(Sender: TObject);
begin
 Application.CreateForm(TfSchool, fSchool);
 fSchool.Mode := 1;
 fSchool.ID := tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag;
 fSchool.Show();
end;

procedure TfMain.bSchoolDelClick(Sender: TObject);
begin
 // Проверяем уверенность юзера в своих намерениях =)
 if (MessageDlg(fData.standarts[9],mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
  fData.DeleteSchool(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag);
 FormShow(self);
end;

procedure TfMain.pmSchoolAddClick(Sender: TObject);
begin
 Application.CreateForm(TfSchool, fSchool);
 fSchool.Mode := 0;
 fSchool.Show();
end;

procedure TfMain.pmSchoolEditClick(Sender: TObject);
begin
 Application.CreateForm(TfSchool, fSchool);
 fSchool.Mode := 1;
 fSchool.ID := tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag;
 fSchool.Show();
end;

procedure TfMain.pmSchoolDelClick(Sender: TObject);
begin
 // Проверяем уверенность юзера в своих намерениях =)
 if (MessageDlg(fData.standarts[9],mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
  fData.DeleteSchool(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag);
 FormShow(self);
end;

procedure TfMain.pmClassAddClick(Sender: TObject);
begin
 Application.CreateForm(TfNewClass, fNewClass);
 fNewClass.Mode := 0;
 fNewClass.School := integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag);
 fNewClass.Show();
end;

procedure TfMain.pmClassEditClick(Sender: TObject);
begin
 Application.CreateForm(TfNewClass, fNewClass);
 fNewClass.Mode := 1;
 fNewClass.ID := integer(sgClassNames.Objects[0, sgClassNames.Row]);
 fNewClass.School := integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag);
 fNewClass.Show();
end;

procedure TfMain.pmClassDelClick(Sender: TObject);
begin
 // Проверяем уверенность юзера в своих намерениях =)
 if (MessageDlg(fData.standarts[9],mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
 begin
  fData.ClassDel(integer(sgClassNames.Objects[0, sgClassNames.Row]));
  PagerChange(Self);
 end;
end;

procedure TfMain.pmClassInClick(Sender: TObject);
begin
 Application.CreateForm(TfClassOb, fClassOb);
 fClassOb.School := integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag);
 fClassOb.Show();
end;

procedure TfMain.pmClassBreakClick(Sender: TObject);
begin
 Application.CreateForm(TfClassRz, fClassRz);
 fClassRz.School := integer(tSchools.AdvOfficeTabs[tSchools.ActiveTabIndex].Tag);
 fClassRz.Show();
end;

procedure TfMain.MnFindClick(Sender: TObject);
begin
 Application.CreateForm(TfFind, fFind);
 fFind.Show();
end;

procedure TfMain.bNextYearClick(Sender: TObject);
begin
 fData.GoNextYear();
 FormShow(self);
end;

procedure TfMain.bPeopMoveClassClick(Sender: TObject);
begin
 if (MessageDlg(fData.standarts[9],mtConfirmation,[mbYes, mbNo], 0) = mrYes) then
 case fData.cUpdate('TB_PEOP', 'CLASS', IntToStr(integer(cbPeopMoveClass.Items.Objects[cbPeopMoveClass.ItemIndex])), 'where ID='+IntToStr(integer(sgClassInfo.Objects[0, sgClassInfo.Row]))) of
  0: sgClassNamesClick(self);
  1: sgClassNamesClick(self);
 end;
end;

procedure TfMain.cbPeopMoveSchoolChange(Sender: TObject);
begin
 // Выбираем классы выбранной школы
 if (integer(cbPeopMoveSchool.Items.Objects[cbPeopMoveSchool.ItemIndex]) <> -1) then
 begin
  case fData.FillClassCb(cbPeopMoveSchoolClass, 'Where (SCHOOL = '+IntToStr(integer(cbPeopMoveSchool.Items.Objects[cbPeopMoveSchool.ItemIndex]))+')') of
   0: if (integer(sgClassInfo.Objects[0, sgClassInfo.Row]) = -1) then bPeopMoveSchool.Enabled := FALSE else bPeopMoveSchool.Enabled := TRUE;
   1: bPeopMoveSchool.Enabled := FALSE;
  end;
 end else bPeopMoveSchool.Enabled := FALSE;
end;

procedure TfMain.bPeopMoveSchoolClick(Sender: TObject);
begin
 if (MessageDlg(fData.standarts[14],mtConfirmation,[mbYes, mbNo], 0) = mrYes) then
 begin
  fData.cDelete('TB_MARKS', 'where PEOP='+IntToStr(integer(sgClassInfo.Objects[0, sgClassInfo.Row])));
  case fData.cUpdate('TB_PEOP', 'CLASS', IntToStr(integer(cbPeopMoveSchoolClass.Items.Objects[cbPeopMoveSchoolClass.ItemIndex])), 'where ID='+IntToStr(integer(sgClassInfo.Objects[0, sgClassInfo.Row]))) of
   0: sgClassNamesClick(self);
   1: sgClassNamesClick(self);
  end;
 end;
end;

procedure TfMain.bMedExcelClick(Sender: TObject);
begin
 fData.ExpMedExcel(sgClassMed);
end;

procedure TfMain.bEndAMedExcelClick(Sender: TObject);
begin
 fData.ExpMedExcel(sgEndMed);
end;

procedure TfMain.MnAdminClick(Sender: TObject);
begin
 Application.CreateForm(TfAdmin, fAdmin);
 fAdmin.ShowModal();
end;

procedure TfMain.SetRules(Filter: byte);
begin
 case fData.Admin of
  // Админ
  0: if (Filter = 0) then
     begin
      MnAdmin.Enabled := TRUE;
      MnAdmin.Visible := TRUE;
      bPredmet.Enabled := TRUE;
      bPredmet.Visible := TRUE;
      bSchoolAdd.Enabled := TRUE;
      bSchoolAdd.Visible := TRUE;
      bSchoolEdit.Enabled := TRUE;
      bSchoolEdit.Visible := TRUE;
      bSchoolDel.Enabled := TRUE;
      bSchoolDel.Visible := TRUE;
      bNextYear.Enabled := TRUE;
      bNextYear.Visible := TRUE;
      bClassAdd.Enabled := TRUE;
      bClassAdd.Visible := TRUE;
      bClassEdit.Enabled := TRUE;
      bClassEdit.Visible := TRUE;
      bClassDel.Enabled := TRUE;
      bClassDel.Visible := TRUE;
      bClassIn.Enabled := TRUE;
      bClassIn.Visible := TRUE;
      bClassBreak.Enabled := TRUE;
      bClassBreak.Visible := TRUE;
      bClassPlan.Enabled := TRUE;
      bClassPlan.Visible := TRUE;
      bTeacherAdd.Enabled := TRUE;
      bTeacherAdd.Visible := TRUE;
      bTeacherDel.Enabled := TRUE;
      bTeacherDel.Visible := TRUE;
      bTeacherSave.Enabled := TRUE;
      bTeacherSave.Visible := TRUE;
      bTeacherMoveSchool.Enabled := TRUE;
      bTeacherMoveSchool.Visible := TRUE;
      bPlanAdd.Enabled := TRUE;
      bPlanAdd.Visible := TRUE;
      bPlanEdit.Enabled := TRUE;
      bPlanEdit.Visible := TRUE;
      bPlanDel.Enabled := TRUE;
      bPlanDel.Visible := TRUE;
      bPlanRAdd.Enabled := TRUE;
      bPlanRAdd.Visible := TRUE;
      bPlanRDel.Enabled := TRUE;
      bPlanRDel.Visible := TRUE;
      bPlanImport.Enabled := TRUE;
      bPlanImport.Visible := TRUE;
      pmSchoolAdd.Enabled := TRUE;
      pmSchoolAdd.Visible := TRUE;
      pmSchoolEdit.Enabled := TRUE;
      pmSchoolEdit.Visible := TRUE;
      pmSchoolDel.Enabled := TRUE;
      pmSchoolDel.Visible := TRUE;
      pmClassAdd.Enabled := TRUE;
      pmClassAdd.Visible := TRUE;
      pmClassEdit.Enabled := TRUE;
      pmClassEdit.Visible := TRUE;
      pmClassDel.Enabled := TRUE;
      pmClassDel.Visible := TRUE;
      pmClassIn.Enabled := TRUE;
      pmClassIn.Visible := TRUE;
      pmClassBreak.Enabled := TRUE;
      pmClassBreak.Visible := TRUE;
      Bevel14.Visible := TRUE;
      Bevel13.Visible := TRUE;
      pUplanNav.Height := 41;
      pClassNav.Height := 47;
     end;
  1: begin
      MnAdmin.Enabled := FALSE;
      MnAdmin.Visible := FALSE;
      bPredmet.Enabled := FALSE;
      bPredmet.Visible := FALSE;
      bSchoolAdd.Enabled := FALSE;
      bSchoolAdd.Visible := FALSE;
      bSchoolEdit.Enabled := FALSE;
      bSchoolEdit.Visible := FALSE;
      bSchoolDel.Enabled := FALSE;
      bSchoolDel.Visible := FALSE;
      bNextYear.Enabled := FALSE;
      bNextYear.Visible := FALSE;
      bClassAdd.Enabled := FALSE;
      bClassAdd.Visible := FALSE;
      bClassEdit.Enabled := FALSE;
      bClassEdit.Visible := FALSE;
      bClassDel.Enabled := FALSE;
      bClassDel.Visible := FALSE;
      bClassIn.Enabled := FALSE;
      bClassIn.Visible := FALSE;
      bClassBreak.Enabled := FALSE;
      bClassBreak.Visible := FALSE;
      bClassPlan.Enabled := FALSE;
      bClassPlan.Visible := FALSE;
      bTeacherAdd.Enabled := FALSE;
      bTeacherAdd.Visible := FALSE;
      bTeacherDel.Enabled := FALSE;
      bTeacherDel.Visible := FALSE;
      bTeacherSave.Enabled := FALSE;
      bTeacherSave.Visible := FALSE;
      bTeacherMoveSchool.Enabled := FALSE;
      bTeacherMoveSchool.Visible := FALSE;
      bPlanAdd.Enabled := FALSE;
      bPlanAdd.Visible := FALSE;
      bPlanEdit.Enabled := FALSE;
      bPlanEdit.Visible := FALSE;
      bPlanDel.Enabled := FALSE;
      bPlanDel.Visible := FALSE;
      bPlanRAdd.Enabled := FALSE;
      bPlanRAdd.Visible := FALSE;
      bPlanRDel.Enabled := FALSE;
      bPlanRDel.Visible := FALSE;
      bPlanImport.Enabled := FALSE;
      bPlanImport.Visible := FALSE;
      pmSchoolAdd.Enabled := FALSE;
      pmSchoolAdd.Visible := FALSE;
      pmSchoolEdit.Enabled := FALSE;
      pmSchoolEdit.Visible := FALSE;
      pmSchoolDel.Enabled := FALSE;
      pmSchoolDel.Visible := FALSE;
      pmClassAdd.Enabled := FALSE;
      pmClassAdd.Visible := FALSE;
      pmClassEdit.Enabled := FALSE;
      pmClassEdit.Visible := FALSE;
      pmClassDel.Enabled := FALSE;
      pmClassDel.Visible := FALSE;
      pmClassIn.Enabled := FALSE;
      pmClassIn.Visible := FALSE;
      pmClassBreak.Enabled := FALSE;
      pmClassBreak.Visible := FALSE;
      Bevel14.Visible := FALSE;
      Bevel13.Visible := FALSE;
      pUplanNav.Height := 0;
      pClassNav.Height := 0;
      cbTeacherMoveSchool.Visible := FALSE;
      LTeacherMoveSchool.Visible := FALSE;
     end;
 end;
end;

procedure TfMain.MnLanguageClick(Sender: TObject);
begin
 if (ODLang.Execute) then
  if (Length(ODLang.FileName) > 0) then fData.ChangeLanguage(ODLang.FileName);
end;

procedure TfMain.bTeacherMoveSchoolClick(Sender: TObject);
begin
 if (fData.cCount('ID', 'TB_PLANCLASS', 'where TEACHER='+IntToStr(integer(sgTeacher.Objects[0, sgTeacher.Row]))) > 0) then
 begin
  fData.ShowError(19);
  exit;
 end;
 fData.cUpdate('TB_TEACHER', 'SCHOOL', IntToStr(integer(cbTeacherMoveSchool.Items.Objects[cbTeacherMoveSchool.ItemIndex])), 'where ID='+IntToStr(integer(sgTeacher.Objects[0, sgTeacher.Row])));
 PagerChange(self);
end;

procedure TfMain.bPeopWordClick(Sender: TObject);
begin
 fData.ExpWord(0, integer(sgClassInfo.Objects[0, sgClassInfo.Row]));
end;

procedure TfMain.bTeacherWordClick(Sender: TObject);
begin
 fData.ExpWord(1, integer(sgTeacher.Objects[0, sgTeacher.Row]));
end;

procedure TfMain.MnExitClick(Sender: TObject);
begin
 Close();
end;

procedure TfMain.MnAboutClick(Sender: TObject);
begin
 Application.CreateForm(TfAbout, fAbout);
 fAbout.Show();
end;

procedure TfMain.MnHelpInfoClick(Sender: TObject);
var s: string;
begin
 s := ExtractFileDir(Application.ExeName)+'\nSchool.chm';
 ShellExecute(Handle,'open',pchar(s),nil,nil,SW_SHOWNORMAL);
end;

procedure TfMain.bPlanDelClick(Sender: TObject);
var plan: integer;
begin
 plan := tPlansClassName.AdvOfficeTabs[tPlansClassName.ActiveTabIndex].Tag;
 // Проверям ведет ли что-то этот учитель
 case fData.ifCount('UPLAN','TB_CLASS','WHERE UPLAN='+IntToStr(plan)) of
  1: fData.ShowError(22,fMain.Handle);
  0: // Проверяем уверенность юзера в своих намерениях =)
     if (MessageDlg(fData.standarts[9],mtConfirmation, [mbYes, mbNo], 0) = mrYes) then
     begin
      case fData.cDelete('TB_PLANCLASS','where UPLAN='+IntToStr(plan)) of
       0: PagerChange(self);
      end;
      case fData.cDelete('TB_UPLAN','where ID='+IntToStr(plan)) of
       0: PagerChange(self);
      end;
     end;
 end;
end;

procedure TfMain.AdvGlassButton1Click(Sender: TObject);
begin
 fData.ExpWord(0, integer(sgEndInfoPeop.Objects[0, sgEndInfoPeop.Row]));
end;

procedure TfMain.AdvSplitter1Moved(Sender: TObject);
begin
 sgClassInfoMarksAll.ColWidths[0] := sgClassInfoMarksAll.Width - 15; 
end;

end.
