unit MyExcel;

interface
  uses Classes,Graphics;

const
//------ Выравнивание по горизонтали -------------------
xlHAlignCenter=-4108;
xlHAlignDistributed=-4117;
xlHAlignJustify=-4130;
xlHAlignLeft=-4131;
xlHAlignRight=-4152;
xlHAlignCenterAcrossSelection=7;
xlHAlignFill=5;
xlHAlignGeneral=1;


//------ Выравнивание по вертикали -------------------
xlVAlignBottom=-4107;
xlVAlignCenter=-4108;
xlVAlignDistributed=-4117;
xlVAlignJustify=-4130;
xlVAlignTop=-4160;


//------- Режимы подчеркивания шрифта -----------------
xlUnderlineStyleNone = -4142;
xlUnderlineStyleSingle = 2;
xlUnderlineStyleDouble = -4119;
xlUnderlineStyleSingleAccounting = 4;
xlUnderlineStyleDoubleAccounting = 5;


//------- Выбор границы ячейки -----------------
xlInsideHorizontal=12;
xlInsideVertical=11;
xlDiagonalDown=5;
xlDiagonalUp=6;
xlEdgeBottom=9;
xlEdgeLeft=7;
xlEdgeRight=10;
xlEdgeTop=8;


//------- Стиль границы ячейки -----------------
xlContinuous=1;
xlDash=-4115;
xlDashDot=4;
xlDashDotDot=5;
xlDot=-4118;
xlDouble=-4119;
xlSlantDashDot=13;
xlLineStyleNone=-4142;


//------- Толщина границы ячейки -----------------
xlHairline=1;
xlThin=2;
xlMedium=-4138;
xlThick=4;


//---------- Тип узора для ячейки ----------------
xlPatternAutomatic=4105;
xlPatternChecker=9;
xlPatternCrissCross=16;
xlPatternDown=-4121;
xlPatternGray16=17;
xlPatternGray25=-4124;
xlPatternGray50=-4125;
xlPatternGray75=-4126;
xlPatternGray8=18;
xlPatternGrid=15;
xlPatternHorizontal=-4128;
xlPatternLightDown=13;
xlPatternLightHorizontal=11;
xlPatternLightUp=14;
xlPatternLightVertical=12;
xlPatternNone=-4142;
xlPatternSemiGray75=10;
xlPatternSolid=1;
xlPatternUp=-4162;
xlPatternVertical=-4166;



//---------- Диалоги ----------------
// печать
xlDialogPrint=8;
xlDialogPrinterSetup=9;
xlDialogPrintPreview=222;

//---------- Ориентация бумаги ------
xlLandscape=2;       // альбомная
xlPortrait=1;        // книжная

//---------- Размер бумаги ----------
xlPaperLetter=1;               //Letter (8-1/2 in. x 11 in.)
xlPaperLetterSmall= 2;         //Letter Small (8-1/2 in. x 11 in.)
xlPaperTabloid= 3;             //Tabloid (11 in. x 17 in.)
xlPaperLedger= 4;              //Ledger (17 in. x 11 in.)
xlPaperLegal= 5;               //Legal (8-1/2 in. x 14 in.)
xlPaperStatement= 6;           //Statement (5-1/2 in. x 8-1/2 in.)
xlPaperExecutive= 7;           //Executive (7-1/2 in. x 10-1/2 in.)
xlPaperA3= 8;                  //A3 (297 mm x 420 mm)
xlPaperA4= 9;                  //A4 (210 mm x 297 mm)
xlPaperA4Small= 10;            //A4 Small (210 mm x 297 mm)
xlPaperA5= 11;                 //A5 (148 mm x 210 mm)
xlPaperB4= 12;                 //B4 (250 mm x 354 mm)
xlPaperB5= 13;                 //B5 (182 mm x 257 mm)
xlPaperFolio= 14;              //Folio (8-1/2 in. x 13 in.)
xlPaperQuarto= 15;             //Quarto (215 mm x 275 mm)
xlPaper10x14= 16;              //10 in. x 14 in.
xlPaper11x17= 17;              //11 in. x 17 in.
xlPaperNote= 18;               //Note (8-1/2 in. x 11 in.)
xlPaperEnvelope9= 19;          //Envelope #9 (3-7/8 in. x 8-7/8 in.)
xlPaperEnvelope10= 20;         //Envelope #10 (4-1/8 in. x 9-1/2 in.)
xlPaperEnvelope11= 21;         //Envelope #11 (4-1/2 in. x 10-3/8 in.)
xlPaperEnvelope12= 22;         //Envelope #12 (4-1/2 in. x 11 in.)
xlPaperEnvelope14= 23;         //Envelope #14 (5 in. x 11-1/2 in.)
xlPaperCsheet= 24;             //C size sheet
xlPaperDsheet= 25;             //D size sheet
xlPaperEsheet= 26;             //E size sheet
xlPaperEnvelopeDL= 27;         //Envelope DL (110 mm x 220 mm)
xlPaperEnvelopeC3= 29;         //Envelope C3 (324 mm x 458 mm)
xlPaperEnvelopeC4= 30;         //Envelope C4 (229 mm x 324 mm)
xlPaperEnvelopeC5= 28;         //Envelope C5 (162 mm x 229 mm)
xlPaperEnvelopeC6= 31;         //Envelope C6 (114 mm x 162 mm)
xlPaperEnvelopeC65= 32;        //Envelope C65 (114 mm x 229 mm)
xlPaperEnvelopeB4= 33;         //Envelope B4 (250 mm x 353 mm)
xlPaperEnvelopeB5= 34;         //Envelope B5 (176 mm x 250 mm)
xlPaperEnvelopeB6= 35;         //Envelope B6 (176 mm x 125 mm)
xlPaperEnvelopeItaly= 36;      //Envelope (110 mm x 230 mm)
xlPaperEnvelopeMonarch= 37;    //Envelope Monarch (3-7/8 in. x 7-1/2 in.)
xlPaperEnvelopePersonal= 38;   //Envelope (3-5/8 in. x 6-1/2 in.)
xlPaperFanfoldUS= 39;          //U.S. Standard Fanfold (14-7/8 in. x 11 in.)
xlPaperFanfoldStdGerman= 40;   //German Standard Fanfold (8-1/2 in. x 12 in.)
xlPaperFanfoldLegalGerman= 41; //German Legal Fanfold (8-1/2 in. x 13 in.)
xlPaperUser= 256;              // User - defined

//----------- Вид документа ----------------------------------
xlNormalView=1;                // Обычный
xlPageBreakPreview=2;          // Разметка страницы



Function  CreateExcel:boolean;
Function  VisibleExcel(visible:boolean):boolean;
Function  AddWorkBook:boolean;
Function  OpenWorkBook(file_:string):boolean;
Function  AddSheet(newsheet:string):boolean;
Function  DeleteSheet(sheet:variant):boolean;
Function  CountSheets:integer;
Function  GetSheets(value:TStrings):boolean;
Function  SelectSheet(sheet:variant):boolean;
Function  SaveWorkBookAs(file_:string):boolean;
Function  CloseWorkBook:boolean;
Function  CloseExcel:boolean;


Function  GetRange(sheet:variant;range:string):variant;
Function  SetRange(sheet:variant;range:string;value_:variant):boolean;
Function  SetColumnWidth(sheet:variant;column:variant;width:real):boolean;
Function  SetRowHeight(sheet:variant;row:variant;height:real):boolean;
Function  GetFormatRange(sheet:variant;range:string):string;
Function  SetFormatRange(sheet:variant;range:string;format:string):boolean;
Function  SetHorizontalAlignment(sheet:variant;range:string;alignment:integer):boolean;
Function  SetVerticalAlignment(sheet:variant;range:string;alignment:integer):boolean;
Function  SetOrientation(sheet:variant;range:string;orientation:integer):boolean;
Function  SetWrapText(sheet:variant;range:string;WrapText:boolean):boolean;
Function  SetMergeCells(sheet:variant;range:string;MergeCells:boolean):boolean;
Function  SetShrinkToFit(sheet:variant;range:string;ShrinkToFit:boolean):boolean;
Function  SetFontRange(sheet:variant;range:string;font:Tfont):boolean;
Function  SetFontRangeEx(sheet:variant;range:string;underlinestyle,colorindex:integer;superscript,subscript:boolean):boolean;
Function  SetBorderRange(sheet:variant;range:string;Edge,LineStyle,Weight,ColorIndex,Color:integer):boolean;
Function  SetPatternRange(sheet:variant;range:string;Pattern,ColorIndex,PatternColorIndex,Color,PatternColor:integer):boolean;

Function  PrintPreview:boolean;
Function  PrintPreviewEx:boolean;
Function  ShowPrintDialog:boolean;
Function  SetBackgroundPicture(sheet:variant;file_:string):boolean;
Function  DisplayGridlines(display:boolean):boolean;
Function  PrintGridlines(sheet:variant;gridline:boolean):boolean;
Function  PageOrientation(sheet:variant;orientation:integer):boolean;
Function  PagePaperSize(sheet:variant;papersize:integer):boolean;
Function  WindowView(view:integer):boolean;
Function  PagePrintArea(sheet:variant;printarea:string):boolean;





implementation


uses ComObj;
 var E:variant;

Function CreateExcel:boolean;
begin
CreateExcel:=true;
try
E:=CreateOleObject('Excel.Application');
except
CreateExcel:=false;
end;
End;


Function VisibleExcel(visible:boolean):boolean;
begin
VisibleExcel:=true;
try
E.visible:=visible;
except
VisibleExcel:=false;
end;
End;


Function AddWorkBook:boolean;
begin
AddWorkBook:=true;
try
E.Workbooks.Add;
except
AddWorkBook:=false;
end;
End;


Function OpenWorkBook(file_:string):boolean;
begin
OpenWorkBook:=true;
try
E.Workbooks.Open(file_);
except
OpenWorkBook:=false;
end;
End;


Function  AddSheet(newsheet:string):boolean;
begin
AddSheet:=true;
try
E.ActiveWorkbook.Sheets.Add;
E.ActiveWorkbook.ActiveSheet.Name:=newsheet;
except
AddSheet:=false;
end;
End;



Function  DeleteSheet(sheet:variant):boolean;
begin
DeleteSheet:=true;
try
E.DisplayAlerts:=False;
E.ActiveWorkbook.Sheets[sheet].Delete;
E.DisplayAlerts:=True;
except
DeleteSheet:=false;
end;
End;


Function  CountSheets:integer;
begin
try
CountSheets:=E.ActiveWorkbook.Sheets.Count;
except
CountSheets:=-1;
end;
End;


Function  GetSheets(value:TStrings):boolean;
 var a_:integer;
begin
GetSheets:=true;
value.Clear;
try
for a_:=1 to E.ActiveWorkbook.Sheets.Count do
    value.Add(E.ActiveWorkbook.Sheets.Item[a_].Name);
except
GetSheets:=false;
value.Clear;
end;
End;



Function  SelectSheet(sheet:variant):boolean;
begin
SelectSheet:=true;
try
E.ActiveWorkbook.Sheets.Item[sheet].Select
except
SelectSheet:=false;
end;
End;



Function SaveWorkBookAs(file_:string):boolean;
begin
SaveWorkBookAs:=true;
try
E.DisplayAlerts:=False;
E.ActiveWorkbook.SaveAs(file_);
E.DisplayAlerts:=True;
except
SaveWorkBookAs:=false;
end;
End;



Function CloseWorkBook:boolean;
begin
CloseWorkBook:=true;
try
E.ActiveWorkbook.Close;
except
CloseWorkBook:=false;
end;
End;



Function CloseExcel:boolean;
begin
CloseExcel:=true;
try
E.Quit;
except
CloseExcel:=false;
end;
End;



//--------------------------------------------------------------------------

Function  SetRange(sheet:variant;range:string;value_:variant):boolean;
begin
SetRange:=true;
try
E.ActiveWorkbook.Sheets.Item[sheet].Range[range]:=value_;
except
SetRange:=false;
end;
End;


Function  GetRange(sheet:variant;range:string):variant;
begin
try
GetRange:= E.ActiveWorkbook.Sheets.Item[sheet].Range[range];
except
GetRange:='';
end;
End;



Function  SetColumnWidth(sheet:variant;column:variant;width:real):boolean;
begin
SetColumnWidth:=true;
try
E.ActiveWorkbook.Sheets.Item[sheet].Columns[column].ColumnWidth:=width;
except
SetColumnWidth:=false;
end;
End;


Function  SetRowHeight(sheet:variant;row:variant;height:real):boolean;
begin
SetRowHeight:=true;
try
E.ActiveWorkbook.Sheets.Item[sheet].Rows[row].RowHeight:=height;
except
SetRowHeight:=false;
end;
End;


//------------ Форматы ячеек -----------------------------------
//    "mmmm yy"
//    "#,##0.00_ ;[Red]-#,##0.00 "
//    "@"
//    "#,##0.00$;[Red]#,##0.00$"

Function  GetFormatRange(sheet:variant;range:string):string;
begin
try
GetFormatRange:=E.ActiveWorkbook.Sheets.Item[sheet].Range[range].NumberFormat;
except
GetFormatRange:='';
end;
End;

Function  SetFormatRange(sheet:variant;range:string;format:string):boolean;
begin
SetFormatRange:=true;
try
E.ActiveWorkbook.Sheets.Item[sheet].Range[range].NumberFormat:=format;
except
SetFormatRange:=false;
end;
End;


//----------------------- Выравнивание ------------------------

// По горизонтали
Function  SetHorizontalAlignment(sheet:variant;range:string;alignment:integer):boolean;
begin
SetHorizontalAlignment:=true;
try
E.ActiveWorkbook.Sheets.Item[sheet].Range[range].HorizontalAlignment:=alignment;
except
SetHorizontalAlignment:=false;
end;
End;

// По вертикали
Function  SetVerticalAlignment(sheet:variant;range:string;alignment:integer):boolean;
begin
SetVerticalAlignment:=true;
try
E.ActiveWorkbook.Sheets.Item[sheet].Range[range].VerticalAlignment:=alignment;
except
SetVerticalAlignment:=false;
end;
End;


//----------------------- Поворот ------------------------
Function  SetOrientation(sheet:variant;range:string;orientation:integer):boolean;
begin
SetOrientation:=true;
try
E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Orientation:=orientation;
except
SetOrientation:=false;
end;
End;


//----------------------- Установка/отмена свойства "перенос по словам" ------------------------
Function  SetWrapText(sheet:variant;range:string;WrapText:boolean):boolean;
begin
SetWrapText:=true;
try
E.ActiveWorkbook.Sheets.Item[sheet].Range[range].WrapText:=WrapText;
except
SetWrapText:=false;
end;
End;


//----------------------- Объединение/отмена объединения ячеек ------------------------
Function  SetMergeCells(sheet:variant;range:string;MergeCells:boolean):boolean;
begin
SetMergeCells:=true;
try
E.ActiveWorkbook.Sheets.Item[sheet].Range[range].MergeCells:=MergeCells;
except
SetMergeCells:=false;
end;
End;


//----------------------- Объединение/отмена объединения ячеек ------------------------
Function  SetShrinkToFit(sheet:variant;range:string;ShrinkToFit:boolean):boolean;
begin
SetShrinkToFit:=true;
try
E.ActiveWorkbook.Sheets.Item[sheet].Range[range].ShrinkToFit:=ShrinkToFit;
except
SetShrinkToFit:=false;
end;
End;

//--------------------------- Изменение шрифта ячейки ---------------------------------

Function  SetFontRange(sheet:variant;range:string;font:Tfont):boolean;
begin
SetFontRange:=true;
try
E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Font.Name:=font.Name;
if fsBold in font.Style      then E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Font.Bold:=True             //           Жирный
                             else E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Font.Bold:=False;           //           Тонкий
if fsItalic in font.Style    then E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Font.Italic:=True           //           Наклонный
                             else E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Font.Italic:=False;         //           Наклонный
E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Font.Size:=font.Size;                                         //           Размер
if fsStrikeOut in font.Style then E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Font.Strikethrough:=True    //           Перечеркнутый
                             else E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Font.Strikethrough:=False;  //           Перечеркнутый
//E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Font.Superscript = False                'Верхний индекс
//E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Font.Subscript = False                  'Нижний индекс
//E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Font.OutlineFont = True                 'Не используется
//E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Font.Shadow = False                     'Не используется
//E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Font.ColorIndex = 41                    'Индекс цвета
if fsUnderline in font.Style then E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Font.Underline:=xlUnderlineStyleSingle // 'Подчеркивание
                             else E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Font.Underline:=xlUnderlineStyleNone;  // 'Подчеркивание
E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Font.Color:=font.Color;                                       //           Цвет
except
SetFontRange:=false;
end;
End;



Function  SetFontRangeEx(sheet:variant;range:string;underlinestyle,colorindex:integer;superscript,subscript:boolean):boolean;
begin
SetFontRangeEx:=true;
try
E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Font.Superscript:=superscript;    //   Верхний индекс
E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Font.Subscript:=subscript;        //   Нижний индекс
E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Font.ColorIndex:=colorindex;      //   Индекс цвета
E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Font.Underline:=underlinestyle;   //   Подчеркивание
except
SetFontRangeEx:=false;
end;
End;



Function  SetBorderRange(sheet:variant;range:string;Edge,LineStyle,Weight,ColorIndex,Color:integer):boolean;
begin
SetBorderRange:=true;
try
E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Borders.item[Edge].LineStyle:=LineStyle;       //   Стиль линн
E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Borders.item[Edge].Weight:=Weight;             //   Толщина линн
if ColorIndex>0 then
   E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Borders.item[Edge].ColorIndex:=ColorIndex   //   Индекс цвета
                else
   E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Borders.item[Edge].Color:=color;            //   Цвет
except
SetBorderRange:=false;
end;
End;



Function  SetPatternRange(sheet:variant;range:string;Pattern,ColorIndex,PatternColorIndex,Color,PatternColor:integer):boolean;
begin
SetPatternRange:=true;
try
E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Interior.Pattern:=Pattern;                                          //   Стиль линн
if ColorIndex>0 then E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Interior.ColorIndex:=ColorIndex                //   Индекс цвета
                else E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Interior.Color:=color;                         //   Цвет
if PatternColorIndex>0
                then E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Interior.PatternColorIndex:=PatternColorIndex  //   Индекс цвета
                else E.ActiveWorkbook.Sheets.Item[sheet].Range[range].Interior.PatternColor:=PatternColor;           //   Цвет
except
SetPatternRange:=false;
end;
End;



Function  PrintPreview:boolean;
begin
PrintPreview:=true;
try
E.ActiveWindow.SelectedSheets.PrintPreview;
except
PrintPreview:=false;
end;
End;



Function  PrintPreviewEx:boolean;
begin
PrintPreviewEx:=true;
try
E.Dialogs.Item[xlDialogPrintPreview].Show;
except
PrintPreviewEx:=false;
end;
End;



Function  ShowPrintDialog:boolean;
begin
ShowPrintDialog:=true;
try
E.Dialogs.Item[xlDialogPrint].Show;
except
ShowPrintDialog:=false;
end;
End;


Function  SetBackgroundPicture(sheet:variant;file_:string):boolean;
begin
SetBackgroundPicture:=true;
try
E.ActiveWorkbook.Sheets.Item[sheet].SetBackgroundPicture(FileName:=file_);
except
SetBackgroundPicture:=false;
end;
End;



Function  DisplayGridlines(display:boolean):boolean;
begin
DisplayGridlines:=true;
try
E.ActiveWindow.DisplayGridlines:=display;
except
DisplayGridlines:=false;
end;
End;


Function  PrintGridlines(sheet:variant;gridline:boolean):boolean;
begin
PrintGridlines:=true;
try
E.ActiveWorkbook.Sheets[sheet].PageSetup.PrintGridlines:=gridline;
except
PrintGridlines:=false;
end;
End;



Function  PageOrientation(sheet:variant;orientation:integer):boolean;
begin
PageOrientation:=true;
try
E.ActiveWorkbook.Sheets[sheet].PageSetup.Orientation:=orientation;
except
PageOrientation:=false;
end;
End;


Function  PagePaperSize(sheet:variant;papersize:integer):boolean;
begin
PagePaperSize:=true;
try
E.ActiveWorkbook.Sheets[sheet].PageSetup.PaperSize:=papersize;
except
PagePaperSize:=false;
end;
End;


Function  WindowView(view:integer):boolean;
begin
WindowView:=true;
try
E.ActiveWindow.View:=view;
except
WindowView:=false;
end;
End;


Function  PagePrintArea(sheet:variant;printarea:string):boolean;
begin
PagePrintArea:=true;
try
E.ActiveWorkbook.Sheets[sheet].PageSetup.PrintArea:=printarea;
except
PagePrintArea:=false;
end;
End;




end.
