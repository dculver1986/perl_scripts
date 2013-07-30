use strict;
use warnings;
use autodie;
use Text::CSV;

my $csv = Text::CSV->new ({ binary => 1 });
my $tsv = Text::CSV->new ({ binary => 1, sep_char => "\t", eol => "\n"});

open my $infh,  "<:encoding(utf8)", "$ARGV[0]";
open my $outfh, ">:encoding(utf8)", "$ARGV[0].txt";

while (my $row = $csv->getline ($infh)) {
$tsv->print ($outfh, $row);
    }
