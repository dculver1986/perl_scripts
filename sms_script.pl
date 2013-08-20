#! /usr/bin/perl
use strict;
use warnings;
use WWW::Mechanize;
use WWW::Twilio::API;
use Data::Dumper;

my $mech = WWW::Mechanize->new();
my @phone = ("7131111111");
$mech->get('http://status.apps.rackspace.com/');
my $content = $mech->content;
my $uh_oh = "field-group-status-icon sprite status-2";


if ($content =~  m/$uh_oh/) {
foreach my $phone_no (@phone){
my $twilio = WWW::Twilio::API->new(AccountSid => '$accountid',
                                AuthToken  => '$authtoken');
my $response = $twilio->POST('SMS/Messages',
                            From => '$twilio',
                            To   => $phone_no,
                            Body => "EMAIL ALERT: There may be issues with email.  Visit http://status.apps.rackspace.com/" );
print "Oh nooooooooooos!!!";
print Dumper($response);
}
}
