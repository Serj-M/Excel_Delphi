unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ComCtrls, Tabnotbk, ExtDlgs;

type
  TForm1 = class(TForm)
    TabbedNotebook1: TTabbedNotebook;
    Button1: TButton;
    ListBox1: TListBox;
    Button2: TButton;
    Button3: TButton;
    Button10: TButton;
    Button4: TButton;
    Memo1: TMemo;
    Button5: TButton;
    Button6: TButton;
    Button7: TButton;
    Button8: TButton;
    CheckBox1: TCheckBox;
    CheckBox2: TCheckBox;
    CheckBox3: TCheckBox;
    Button9: TButton;
    FontDialog1: TFontDialog;
    Button11: TButton;
    Button12: TButton;
    Button13: TButton;
    Button14: TButton;
    Button15: TButton;
    Button16: TButton;
    Button17: TButton;
    Button18: TButton;
    Button19: TButton;
    Button20: TButton;
    OpenPictureDialog1: TOpenPictureDialog;
    CheckBox4: TCheckBox;
    CheckBox5: TCheckBox;
    ComboBox1: TComboBox;
    ComboBox2: TComboBox;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button10Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure CheckBox2Click(Sender: TObject);
    procedure CheckBox3Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure Button11Click(Sender: TObject);
    procedure Button12Click(Sender: TObject);
    procedure Button13Click(Sender: TObject);
    procedure Button14Click(Sender: TObject);
    procedure Button15Click(Sender: TObject);
    procedure Button16Click(Sender: TObject);
    procedure Button17Click(Sender: TObject);
    procedure Button18Click(Sender: TObject);
    procedure Button19Click(Sender: TObject);
    procedure Button20Click(Sender: TObject);
    procedure CheckBox4Click(Sender: TObject);
    procedure CheckBox5Click(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
    procedure ComboBox2Change(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.DFM}
   uses MyExcel;

procedure TForm1.Button1Click(Sender: TObject);
 var a_:integer;
   rng_:string;
begin
if not CreateExcel then exit;
messagebox(handle,'','Запускаем Excel.',0);
VisibleExcel(true);
messagebox(handle,'','Отобразили Excel на экране.',0);
if AddWorkBook then begin
messagebox(handle,'','Создали новую книгу.',0);
AddSheet('Новый лист');
messagebox(handle,'','Добавили новый лист.',0);
DeleteSheet(2);
messagebox(handle,'','Удалили лист №2.',0);
GetSheets(ListBox1.Items);
messagebox(handle,'','Получили список листов!',0);
for a_:=1 to CountSheets do begin
    ListBox1.ItemIndex:=a_-1;
    SelectSheet(ListBox1.Items.Strings[a_-1]);
    messagebox(handle,'',pchar('Выбираем лист '+ListBox1.Items.Strings[a_-1]+' !'),0);
    end;

SetRange(3,'A1',234.45);
rng_:=GetRange(3,'A1');
messagebox(handle,pchar(rng_),'Записываем и читаем ячейку.',0);

SaveWorkBookAs('c:\1.xls');
messagebox(handle,'','Сохранили книгу как "c:\1.xls".',0);
CloseWorkBook;
messagebox(handle,'','Закрыли книгу "c:\1.xls".',0);
                    end;
CloseExcel;
end;

procedure TForm1.Button2Click(Sender: TObject);
begin
if not CreateExcel then exit;
VisibleExcel(true);
if AddWorkBook then begin
   SelectSheet(1);
   Button3.Enabled:=true;
   Button10.Enabled:=true;
   end;
end;


procedure TForm1.Button3Click(Sender: TObject);
 var a_:integer;
   rng_:string;
begin
for a_:=1 to 100 do begin
    SetRange(1,'A'+inttostr(a_),a_);
    SetRange(1,'B'+inttostr(a_),234.45);
    SetRange(1,'C'+inttostr(a_),random(1000));
    end;
Button4.Enabled:=true;
end;

procedure TForm1.Button10Click(Sender: TObject);
begin
SaveWorkBookAs('c:\1.xls');
messagebox(handle,'','Сохранили книгу как "c:\1.xls".',0);
CloseWorkBook;
messagebox(handle,'','Закрыли книгу "c:\1.xls".',0);
CloseExcel;
end;

procedure TForm1.Button4Click(Sender: TObject);
 var a_:integer;
   eee_,bbb_:string;
begin
for a_:=1 to 100 do begin
    eee_:=GetRange(1,'A'+inttostr(a_));
    eee_:=eee_+'/';
    bbb_:=GetRange(1,'B'+inttostr(a_));
    eee_:=eee_+bbb_;
    eee_:=eee_+'/';
    bbb_:=GetRange(1,'C'+inttostr(a_));
    eee_:=eee_+bbb_+'      '+GetFormatRange(1,'C'+inttostr(a_));
    Memo1.Lines.Add(eee_);
    end;
Button5.Enabled:=true;
end;



procedure TForm1.Button5Click(Sender: TObject);
begin
SetColumnWidth(1,2,50.25);
SetRowHeight(1,2,30.10);
Button6.Enabled:=true;
end;

procedure TForm1.Button6Click(Sender: TObject);
begin
messagebox(handle,'','Изменение формата данных ячейки A1',0);
SetFormatRange(1,'A1:C5','0,000');
Button7.Enabled:=true;
end;

procedure TForm1.Button7Click(Sender: TObject);
begin
SetHorizontalAlignment(1,'A1',xlHAlignCenter);
SetVerticalAlignment(1,'A1',xlVAlignBottom);
Button8.Enabled:=true;
end;

procedure TForm1.Button8Click(Sender: TObject);
begin
SetOrientation(1,'A1:C5',-90);
CheckBox1.Enabled:=true;
CheckBox2.Enabled:=true;
CheckBox3.Enabled:=true;
end;

procedure TForm1.CheckBox1Click(Sender: TObject);
begin
SetWrapText(1,'A1',CheckBox1.Checked);
end;

procedure TForm1.CheckBox2Click(Sender: TObject);
begin
SetMergeCells(1,'A1:B2',CheckBox2.Checked);
end;

procedure TForm1.CheckBox3Click(Sender: TObject);
begin
SetShrinkToFit(1,'A1',CheckBox3.Checked);
end;

procedure TForm1.Button9Click(Sender: TObject);
begin
if FontDialog1.Execute then SetFontRange(1,'A1',FontDialog1.Font);
end;

procedure TForm1.Button11Click(Sender: TObject);
begin
SetFontRangeEx(1,'A1',xlUnderlineStyleDouble,1,true,false);
end;

procedure TForm1.Button12Click(Sender: TObject);
begin
SetBorderRange(1,'B2',xlDiagonalDown,xlDashDot,xlThin,0,rgb(50,150,230));
SetBorderRange(1,'B2',xlDiagonalUp,xlDashDot,xlThin,0,rgb(50,150,230));
SetBorderRange(1,'B2',xlEdgeBottom,xlDouble,xlThin,0,rgb(0,200,230));
SetBorderRange(1,'B2',xlEdgeLeft,xlDouble,xlThin,0,rgb(0,200,230));
SetBorderRange(1,'B2',xlEdgeRight,xlDouble,xlThin,0,rgb(0,200,230));
SetBorderRange(1,'B2',xlEdgeTop,xlDouble,xlThin,0,rgb(0,200,230));
end;

procedure TForm1.Button13Click(Sender: TObject);
begin
SetPatternRange(1,'A1:C4',xlPatternGray8,-1,-1,rgb(100,25,200),rgb(0,100,200));
end;

procedure TForm1.Button14Click(Sender: TObject);
begin
if not CreateExcel then exit;
VisibleExcel(true);
if AddWorkBook then begin
   SelectSheet(1);
   Button15.Enabled:=true;
   Button16.Enabled:=true;
   end;
ComboBox1.ItemIndex:=0;
ComboBox2.ItemIndex:=0;
end;

procedure TForm1.Button15Click(Sender: TObject);
 var a_:integer;
   rng_:string;
begin
for a_:=1 to 20 do begin
    SetRange(1,'A'+inttostr(a_),a_);
    SetRange(1,'B'+inttostr(a_),234.45);
    SetRange(1,'C'+inttostr(a_),random(1000));
    end;
end;


procedure TForm1.Button16Click(Sender: TObject);
begin
SaveWorkBookAs('c:\1.xls');
messagebox(handle,'','Сохранили книгу как "c:\1.xls".',0);
CloseWorkBook;
messagebox(handle,'','Закрыли книгу "c:\1.xls".',0);
CloseExcel;
end;

procedure TForm1.Button17Click(Sender: TObject);
begin
PrintPreview;
end;

procedure TForm1.Button18Click(Sender: TObject);
begin
ShowPrintDialog;
end;

procedure TForm1.Button19Click(Sender: TObject);
begin
PrintPreviewEx;
end;

procedure TForm1.Button20Click(Sender: TObject);
begin
if not OpenPictureDialog1.Execute then exit;
SetBackgroundPicture(1,OpenPictureDialog1.FileName);
end;

procedure TForm1.CheckBox4Click(Sender: TObject);
begin
DisplayGridlines(CheckBox4.Checked);
end;

procedure TForm1.CheckBox5Click(Sender: TObject);
begin
PrintGridlines(1,CheckBox5.Checked);
end;

procedure TForm1.ComboBox1Change(Sender: TObject);
begin
case ComboBox1.ItemIndex of
0:PageOrientation(1,xlPortrait);
1:PageOrientation(1,xlLandscape);
end;
end;

procedure TForm1.ComboBox2Change(Sender: TObject);
begin
case ComboBox2.ItemIndex of
0:WindowView(xlNormalView);
1:WindowView(xlPageBreakPreview);
end;
end;

end.
