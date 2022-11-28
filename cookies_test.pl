#!/usr/bin/env perl

use strict;
use warnings;
use utf8;
use v5.14;
use Data::Dumper;
use Excel::Writer::XLSX;

my $cookies = { 
              Hungary => [ 'bejgli', 'hólabda' ],
              Scandinavia => [ 'Pepparkakor', 'Ruiskakut', 'Lussekatter', 'Krumkake' ],
              Netherlands => [ 'Spekulaas', 'Kerstkranjes'],
              Germany => [ 'Lebkuchen', 'Stollen' ],
              Switzerland => [ 'Basler Läckerli', 'Brunsli' ],
              Italy => [ 'Panettone', 'Pignoli', 'Cannoli' ],
              Spain => [ 'Polvorones', 'Mantecados', 'Turrón' ],
              Greece => [ 'Melomakarona' ],
              };
              
  my %props = ( color => 'red',
              size => '20',
              bg_color => 'green',
              pattern => 18,
              align => 'center',
             );

#Create excel file               
my $merry = Excel::Writer::XLSX->new( 'xmas.xlsx' );
#create the workbooks
my $xmas = $merry->add_worksheet('Merry');
my $baking_sheet = $merry->add_worksheet('Xmas');
my $stats = $merry->add_worksheet('Everybody');

#Add christmas style formatting to the Xmas worksheet.
my $format = $merry->add_format( %props );
#write welcome text with formatting to top left cell
$xmas->write( 'A1', 'Merry Christmas!', $format);
#Set the column width.
$xmas->set_column('A:A', 30); 

my $format2 = $merry->add_format();

$format2->set_color('red');
$format2->set_bold();
$format2->set_size( 12 );
$format2->set_bg_color( 'green' );
$format2->set_align('center');
$format2->set_border( 4 );

my $format3 = $merry->add_format( border => 4,);

#Sort the countries alphabetically
my @countries = sort keys %$cookies;

#Add the countries to the baking sheet, one country per column.
$baking_sheet->set_column( 0, scalar keys %$cookies, 15);
$baking_sheet->write( 'A1', \@countries, $format2 );

my $i = 0;
for ( sort keys %$cookies ) {
  #Place Your cookies to the baking sheet
  $baking_sheet->write( 1, $i, [ $cookies->{$_} ], $format3 );
  $i++;
}

#write some statistics about how many types of cookies we have per country
$stats->write( 'A1', [\@countries], $format2 );


$merry->close();