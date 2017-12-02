#!/usr/bin/perl -w

use DBI;
use Excel::Writer::XLSX;

my $query = $ARGV[0];
my $dbname = '';
my $host = 'localhost';
my $port = '5432';
my $username = '';
my $password = '';

my $dbh = DBI->connect("dbi:Pg:dbname=$dbname;host=$host;port=$port",
  $username,
  $password,
  {AutoCommit => 0, RaiseError => 1, PrintError => 0}
);

my $sth = $dbh->prepare($query);

$sth->execute or die $sth->errstr;


# Create a new Excel workbook
my $workbook = Excel::Writer::XLSX->new( '/tmp/perl.xlsx' );

# Add a worksheet
$worksheet = $workbook->add_worksheet();

#  Add and define a format
$format = $workbook->add_format();
$format->set_bold();
$format->set_color( 'red' );
$format->set_align( 'center' );

# Write a formatted and unformatted string, row and column notation.
$col = $row = 0;
$worksheet->write( $row, $col, 'Hi Excel!', $format );
$worksheet->write( 1, $col, $query );

$sth->finish;
$dbh->disconnect;
