use strict;
use warnings;
# Loading  the Text::CSV_XS and Spreadsheet::ParseXLSX modules
use Text::CSV_XS;
use Spreadsheet::ParseXLSX;
# Creating a new csv file
my $csv = Text::CSV_XS->new({  });
# Creating a new Spreadsheet
my $parser = Spreadsheet::ParseXLSX->new();
#giving name of the input that need to be read
my $excel_file = 'input.xlsx';
#converting file into readable by perl(parsing)
my $excel_workbook = $parser->parse($excel_file);

if (!defined $excel_workbook) {
    die $parser->error(), ".\n";
}
#getting list of work sheets
my @worksheets = $excel_workbook->worksheets();
# Initializing array
my @data;
#this loop will read all data in excel and arranges in data variable in same matrix form
for my $worksheet (@worksheets) {
    my ($row_min, $row_max) = $worksheet->row_range();
    my ($col_min, $col_max) = $worksheet->col_range();
    #for each row and column(element by element)
    for my $col ($col_min..$col_max) {
        my @column_data;
        for my $row ($row_min..$row_max) {
            my $cell = $worksheet->get_cell($row, $col);
        #pushes only if value is there
            if ($cell) {
                push @column_data, $cell->value();
            }
        }
        push @data, \@column_data;#storing in data variable
    }
}
#declaring corner variables
my @corner;
my @enable;
my @moduleName;
my @vdd12;
my @temperature;
#pushing strings
push @corner, "Corner";
push @enable, "Enable";
push @moduleName, $data[3][0];
push @vdd12, "vdK";
push @temperature, "Temperature";
#there three loops as we require combinations of three variable values
for my $l (0..$#{$data[0]}) {
    for my $m (0..$#{$data[1]}) {
        for my $n (0..$#{$data[2]}) {
            next if (!$data[0][$l] ||  !$data[2][$m] || !$data[2][$n]);
            #checking whether value exists for each variable
            my $modName = "$data[0][$l]";
            my $vdd = "$data[1][$m]";
            my $tmp = "$data[2][$n]";
            #pushing in series like vector
            push @moduleName, $modName;
            push @vdd12, $vdd;
            push @temperature, $tmp;
        }
    }
}
#assigning corner name and mentioning true
for my $i (0..$#temperature) {
    push @corner, "C$i";
    push @enable, "t";
}
#generating csv file
my $csv_file = 'OUT.csv';
# opening the output file for writing
open(my $fh, '>', $csv_file) or die "Could not open file '$csv_file' $!";
print $fh join(",", @corner) . "\n";#appending each vector elements in the format
print $fh join(",", @enable) . "\n";
print $fh join(",", @vdd12) . "\n";
print $fh join(",", @temperature) . "\n";
print $fh join(",", @moduleName) . "\n";


close $fh;
