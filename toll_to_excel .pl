#! /usr/bin/perl
use strict;
use warnings;

use DateTime;
use Date::Format;
use DBI;
use Spreadsheet::WriteExcel;

#author @danielc

my $dt = DateTime->today();
my $start_date = $dt->set_day(1)->strftime("%Y%m%d");

#established current month to calculate last day of current month to insert into query language
my $current_month = DateTime->today();

$current_month = DateTime->last_day_of_month(
    month => $current_month->month,
    year  => $current_month->year,
 )->strftime("%Y%m%d");


my $database = "database";
my $host     = "database.local";
my $user     = "me";
my $pw       = "";

my $dbh = DBI->connect("dbi:mysql:$database:$host", $user, $pw,) or die "Database connection FAILED: $DBI::errstr";



my $sth = $dbh->prepare(
"drop TABLE workflow.temp_tca_envelopes");
$sth->execute();


$sth = $dbh->prepare(
"query language here (removed for privacy)"
);
$sth->execute();

$sth = $dbh->prepare(
"ALTER TABLE workflow.temp_tca_envelopes ADD PRIMARY KEY (first_document)"
);

$sth->execute();

my $sth2 = $dbh->prepare(
"query language here (removed for privacy)"
);
$sth2->execute();


my $xls = "tca_report.xls";

my $xWB = Spreadsheet::WriteExcel->new($xls);
my $xWS = $xWB->add_worksheet('TCA_Monthly');

my $format = $xWB->add_format();
$format->set_bold();

$xWS->set_row(0, undef, $format);
$xWS->set_column('A:A', 15);
$xWS->set_column('B:B', 28);
$xWS->set_column('C:C', 28);
$xWS->set_column('D:D', 28);

my $R=0; my $C=0;
$xWS->write($R, $C++, $_) for @{$sth2->{NAME}};

while (my $ar=$sth2->fetchrow_arrayref) {
     ++$R; $C=0;
         $xWS->write($R, $C++, $_) for @$ar;
}
#get row total and for sums
my $R_count = scalar($R)+ 1;
my $totals_R = $R_count + 3;

$xWS->write("A".$totals_R,'Total', $format);
$xWS->write_formula("B".$totals_R,"SUM(B2:B$R_count)",$format);
$xWS->write_formula("C".$totals_R,"SUM(C2:C$R_count)",$format);
$xWS->write_formula("D".$totals_R,"SUM(D2:D$R_count)",$format);


my $sth3 = $dbh->prepare(
"query language here (removed for privacy)"
);
$sth3->execute();

# add second worksheet
my $xWS2 = $xWB->add_worksheet('TCA_Daily');

$xWS2->set_row(0, undef, $format);
$xWS2->set_column('A:A', 10);
$xWS2->set_column('B:B', 10);
$xWS2->set_column('C:C', 12);
$xWS2->set_column('D:D', 25);
$xWS2->set_column('E:E', 22);
$xWS2->set_column('F:F', 22);
$xWS2->set_column('G:G', 10);
$xWS2->write('G1', 'Mail Date', $format);

my $row = 0; my $col = 0;
$xWS2->write($row, $col++, $_) for @{$sth3->{NAME}};

while (my $array_ref = $sth3->fetchrow_arrayref) {
    ++$row, $col = 0;
        $xWS2->write($row, $col++, $_) for @$array_ref;
}
#get row total and for sums
my $row_count = scalar($row) + 1;
my $totals_row = $row_count + 3;

$xWS2->write("A".$totals_row,'Total', $format);
$xWS2->write_formula("D".$totals_row,"SUM(D2:D$row_count)",$format);
$xWS2->write_formula("E".$totals_row,"SUM(E2:E$row_count)",$format);
$xWS2->write_formula("F".$totals_row,"SUM(F2:F$row_count)",$format);


$xWB->close();
$sth->finish();
$sth2->finish();
$sth3->finish();
$dbh->disconnect;




