
' https://github.com/ReneNyffenegger/development_misc/blob/master/vba/runVBAFilesInOffice.vbs
'
'
'   runVBAFilesInOffice.vbs -excel csvToDiagram -c Go %CD%\data.csv

option explicit

public sub Go(csv_file_name as string)

    dim this_workbook      as workbook
    dim sheet_csv_data     as workSheet
    dim sheet_diagram      as chart
    dim qt_csv_data        as queryTable

    set this_workbook      = application.workbooks(1)

    set sheet_csv_data   = this_workbook.workSheets(1)
    sheet_csv_data.name  ="csv data" ' not strictly necessary

    set sheet_diagram    = this_workbook.sheets.add (type := xlChart)
    sheet_diagram.name   ="diagram"  ' not strictly necessary

    set qt_csv_data = CSV_import(      _
          csv_file_name              , _
          array (                      _
            xlTextFormat   , _
            xlGeneralFormat, _
            xlGeneralFormat, _
            xlGeneralFormat, _
            xlGeneralFormat, _
            xlGeneralFormat, _
            xlGeneralFormat, _
            xlGeneralFormat, _
            xlGeneralFormat)         , _
          sheet_csv_data             , _
          sheet_csv_data.range("a1")   _
    )


    Call sheet_diagram.setSourceData(source := qt_csv_data.resultRange) ' or     := range(qt_csv_data.name)

    call setPageUp (sheet_diagram)

    call formatChart (sheet_diagram)

    this_workbook.saved = true

end sub'



private function CSV_import(               _
    csv_file_name       as string     , _
    column_data_types   as variant    , _
    dest_sheet          as worksheet  , _
    csv_range_top_left  as range) as queryTable ' {

  ' TODO: can column_data_types' data type be declared more concisely as just «variant».

    dim dest as range
    dim qt   as queryTable


    set qt   = dest_sheet.QueryTables.Add ( Connection  := "TEXT;" & csv_file_name, _
                                            Destination := csv_range_top_left )


'   qt.name              ="csv_data"    ' not strictly necessary
    qt.refreshOnFileOpen = false
    qt.adjustColumnWidth = true
    qt.textFilePlatform  = 850 ' MS-Dos ?
    qt.textFileStartRow  = 1

    qt.textFileParseType     = xlDelimited
    qt.textFileTextQualifier = xlTextQualifierDoubleQuote

 ' --------------------------------------------

'   qt.textFileConsecutiveDelimiter = false
'   qt.textFileTabDelimiter         = true
    qt.textFileSemicolonDelimiter   = true
'   qt.textFileCommaDelimiter       = false
'   qt.textFileSpaceDelimiter       = false

 ' --------------------------------------------

   ' Define the «data type» of the imported columns.

    qt.textFileColumnDataTypes = column_data_types

    qt.refresh

    set CSV_import = qt

end function ' }

private sub setPageUp(sh as variant) ' {

  dim ps as pageSetup

  set ps = sh.pageSetup

  ps.leftMargin   = application.centimetersToPoints(0.5)
  ps.rightMargin  = application.centimetersToPoints(0.5)
  ps.topMargin    = application.centimetersToPoints(0.5)
  ps.bottomMargin = application.centimetersToPoints(0.5)

  ps.headerMargin = application.centimetersToPoints( 0 )
  ps.footerMargin = application.centimetersToPoints( 0 )

end sub ' }

private sub formatChart(ch as chart) ' {


  dim columnName as string

  dim s as string

  ch.chartType = xlLine

  ch.plotArea.top    =   9
  ch.plotArea.left   =  45
  ch.plotArea.width  = 748
  ch.plotArea.height = 480


  call formatLegend (ch)


  call formatSeries(ch, "FP14"    , 2, 255, 127,   0)

  call formatSeries(ch, "TSE10"   , 3,  30,  90, 120)
  call formatSeries(ch, "TSE11"   , 1,  70,  80, 150)
  call formatSeries(ch, "TSE12"   , 1, 110,  70, 180)
  call formatSeries(ch, "TSE13"   , 1, 130,  20, 210)

  call formatSeries(ch, "SCE10"   , 3,  40, 150, 100)
  call formatSeries(ch, "SCE11"   , 1,  50, 200,  30)
  call formatSeries(ch, "SCE12"   , 1,  60, 250,  70)

end sub ' }

private sub formatLegend(ch as chart) ' {

  dim leg as legend

  set leg = ch.legend

  leg.includeInLayout = false

  leg.format.fill.foreColor.objectThemeColor = msoThemeColorBackground1
  leg.format.fill.transparency = 0.3
  leg.format.fill.solid

  leg.top    =   23.5
  leg.left   =  120.7
  leg.height =  180.6
  leg.width  =   66.8

end sub ' }

private sub formatSeries(ch as chart, seriesName as string, width as double, r as integer, g as integer, b as integer) ' {

   dim ser as series

   '  cstr()?
   '  See http://stackoverflow.com/questions/12620239/what-is-the-difference-between-string-variable-and-cstrstring-variable
   set ser = ch.seriesCollection.item(  cstr(seriesName) )

   ser.format.line.weight        = width
   ser.format.line.foreColor.rgb = rgb(r, g, b)


end sub ' }

