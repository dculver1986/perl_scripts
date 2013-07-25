use warnings;
use WWW::Mechanize;
use Spreadsheet::WriteExcel;
use Text::CSV_XS;
use DateTime;
use Date::Format;
use MIME::Entity;
#Author danielc. It may be best to create a separate folder for each start/end date reports desired
#the script creates a csv file and xls file


#This block of code follows from the CAPS homepage to the Account Inquiries link
my $mech = WWW::Mechanize->new();
my $url = "http://site.to.crawl.com";
    $mech->get($url);
    $mech->follow_link( url => 'https://site.to.crawl.com');
if ($mech->success()){
print "Successful Connection. Retrieving Data\n";
} else {
  print "Not a successful connection\n"; }




# The next 2 blocks of code set start and end dates required for the report
#my $start_date = DateTime->today->subtract( months => 1)->strftime("%m%d%y");

$dt = DateTime->today();
my $mnth = $dt->subtract(months => 1);
my $start_date = $mnth->clone->set_day(1)->strftime("%m%d%y");



my $end_date = DateTime->today->subtract(months => 1);
$end_date = DateTime->last_day_of_month(
    month => $end_date->month,
    year  => $end_date->year,
 )->strftime("%m%d%y");
#print $end_date;



#This block of code is intended to fill in the required forms
$mech->get("https://site.to.crawl.com");
my $usr = "username";
my $pw = "pw";
$mech->form_number(1);
$mech->field( "capsn", $usr);
$mech->form_number(2);
$mech->field("capsp", $pw);
$mech->form_number(3);
$mech->field( "startdate", $start_date);
$mech->form_number(4);
$mech->field( "enddate", $end_date);
$mech->form_number(5);
$mech->select("format", "2");
$mech->click();

#this block of code matches the lines of the list that contain AZ
my $match = "AZ";
my $content = $mech->content; 
my @lines = split /^/, $content;
my @keepers = grep {/\Q$match\E/}  @lines;
print @keepers;







# change start and end date formats
$dt2 = DateTime->today();
my $m = $dt2->subtract(months => 1);
my $start= $m->clone->set_day(1)->strftime("%Y%m%d");



my $end = DateTime->today->subtract(months => 1);
$end = DateTime->last_day_of_month(
     month => $end->month,
         year  => $end->year,
          )->strftime("%Y%m%d");
          


#This code creates a csv file to be converted to xls
open(FH, ">/path/to/caps/csv_"."$start"."_"."$end.csv") or die "$!";
print FH @keepers;
close(FH);



#open csv file for reading
open (FH, "</path/to/csv_"."$start"."_"."$end.csv") or die "Cannot open file: $!\n";


#create excel file
my $workbook = Spreadsheet::WriteExcel->new("/path/to/report_"."$start"."_"."$end.xls");
my $worksheet = $workbook->add_worksheet('CAPS');

#set format for column headers
my $format = $workbook->add_format();
$format->set_bold();
#set currency format
my $trans_format = $workbook->add_format();
$trans_format->set_num_format('$0.00');



#set column format 
$worksheet->set_column('A:C',20);
$worksheet->set_column('D:I', 10);
$worksheet->set_column('I:I', undef, $trans_format);
#write the column headers
$worksheet->write(0,0, 'TOTAL',$format);
$worksheet->write(1,0, 'transaction_no', $format);
$worksheet->write(1,1, 'eff_date',$format);
$worksheet->write(1,2, 'city_name',$format);
$worksheet->write(1,3, 'state',$format);
$worksheet->write(1,4, 'permit_no',$format);
$worksheet->write(1,5, 'permit_type',$format);
$worksheet->write(1,6, 'class',$format);
$worksheet->write(1,7, 'pieces',$format);
$worksheet->write(1,8, 'trans_amt',$format);



#get the lines from csv file and write them to report xls file  
my $row = 2;
while (<FH>) {
    chomp;
    @keepers = split /,/, $_, 10;
    pop @keepers; # only get first 9 columns
    $worksheet->write_row($row, 0, \@keepers);
    $row++;
	
}   


 
$worksheet->write_formula('I1', "=SUM (I3:I$row)");
$worksheet->set_column('I:I', undef, $trans_format);


#close workbook
close(FH);

$workbook->close();
