use warnings;
use strict;

use Win32::OLE;
use Win32::OLE::Const 'Microsoft.Excel';

my ($excel, $workbook, $sheet_data, $sheet_diagram) = prepare_excel();

$sheet_data -> Cells(1, 1) -> {Value}= "x";
$sheet_data -> Cells(1, 2) -> {Value}= "sin(x)";

my $row = 1;
my @x_values;
for (my $x=-3.2; $x<=3.2; $x+=0.05) {
  
  $row++;

  $sheet_data->Cells($row, 1) -> {Value} = sin($x);

  push @x_values, $x;

}

# $sheet_diagram -> {ChartType} =  xlLine;
  $sheet_diagram -> {ChartType} =  xlXYScatterSmoothNoMarkers;

my $data_range = $sheet_data -> Range($sheet_data->Cells(1, 1),
                                      $sheet_data->Cells($row, 1));

my $series_collection = $sheet_diagram -> SeriesCollection;
$series_collection -> Add ({Source => $data_range});

my $series_sin_x = $series_collection -> Item(1);
$series_sin_x -> {XValues} = \@x_values;

$workbook->{Saved}=1;

sub prepare_excel { # {{{

    my $excel = CreateObject Win32::OLE 'Excel.Application' or die;
    $excel -> {Visible} = 1;

    my $workbook  = $excel -> Workbooks -> Add;

    my $sheet_data = $workbook -> WorkSheets(1);
    $sheet_data -> {Name} = "data"; # not striclty necessary

    my $sheet_diagram = $workbook -> Sheets -> Add ({type=>xlChart});
    $sheet_diagram -> {Name} = "diagram"; # not strictly necessary

    return ($excel, $workbook, $sheet_data, $sheet_diagram);
    
} # }}}
