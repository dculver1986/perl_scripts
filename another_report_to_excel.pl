#! /usr/bin/perl

use DBI;
use MIME::Entity;
use Spreadsheet::WriteExcel;
use DateTime;
use Date::Format;


# @danielc

$database = "database";
$host     = "db.local";
$user     = "me";
$pw       = "";

#connect to database
$dbh = DBI->connect("dbi:mysql:$database:$host", $user, $pw,) or die "Database connection FAILED: $DBI::errstr";

#IMB query
my $query = <<THISSQL;
query language removed (for privacy)

THISSQL




$sth = $dbh->prepare($query);
$sth->execute();

# count number of rows
my $rv = $sth->rows;



#get current date
my $time = localtime(time);
my $date = time2str("%m%d%y", time);


#create new spreadsheet
my $xWB = Spreadsheet::WriteExcel->new("/path/to/imb_report_"."$date".".xls");
my $xWS = $xWB->add_worksheet('IMB');

my $format = $xWB->add_format();
$format->set_bold();

$xWS->set_column('A:A', 12);
$xWS->set_column('B:D',8);
$xWS->set_column('F:K', 14,$format); 




# Add column headers for first worksheet
my $R=0; my $C=0;
$xWS->write($R, $C++, $_) for @{$sth->{NAME}};



# Read the query results and write them into the spreadsheet
while (my $ar=$sth->fetchrow_arrayref) {
    ++$R; $C=0;
           $xWS->write($R, $C++, $_) for @$ar;
        }

# create new columns and write their data/formulas
my $lines = $rv + 1;
$xWS->write(0, 5, '# of 81 STID'); 
$xWS->write_formula(1, 5, "=COUNTIF(B2:B$lines, 81)");
                
$xWS->write(0,10,'# of Envelopes');
$xWS->write_formula(1,10, "SUM(D2:D$lines)");
$xWS->write(0,6, '# of 140 STID');
$xWS->write_formula(1,6, "COUNTIF(B2:B$lines, 140)");
$xWS->write(0,7, '# of 240 STID');
$xWS->write_formula(1,7, "COUNTIF(B2:B$lines, 240)");
$xWS->write(0,8, '# of 700 STID');
$xWS->write_formula(1,8, "COUNTIF(B2:B$lines, 700)");
$xWS->write(0,9, '# of 310 STID');
$xWS->write_formula(1,9, "COUNTIF(B2:B$lines, 310)");


$xWB->close();



#create the email
my $top = MIME::Entity->build(Type=>"multipart/mixed",
From=> "production\@work.com",
To=> "someone\@here.com",
Subject=> "IMB_Report_"."$date");


# attach xls file
$top->attach(
              Type => 'application/vnd.ms-excel',
              Encoding => 'base64',
              Path => "/path/to/imb_report_"."$date".".xls",
              Filename => "imb_report_"."$date".".xls"
 );
#send the mail
open MAIL, "|/usr/sbin/sendmail -t";
$top->print(\*MAIL);
close MAIL;



$sth->finish();
$dbh->disconnect();




