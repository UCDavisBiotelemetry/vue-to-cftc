#!/usr/bin/perl 

# feed this the name of a VUE csv file and it will reformat it.  The 
# Date::Calc module is used to convert UTC to PST time.
# Steve Lindley, 25 Mar 2008.
# Modified on 12/15/11 by CJM to also output codespace for the new Detections_all table format
# Modified on 12/27/11 by CJM to no longer output sync and bin
# Clarified on 20140904 by Matt Pagel
# Modified for VUE version 2 CSV output on 20141014 Matt Pagel
# Various error handling for atypical VRL -> CSV output added 20141024 MP

# VUE 1.0 CSV output (14 columns, 13 commas)
# DateTime,Code Space,ID,Sensor 1,Units 1,Sensor 2,Units 2,Transmitter Name,Transmitter S/N,Receiver Name,Receiver S/N,Station Name,Station Latitude,Station Longitude

# VUE 2.0 CSV output
# Date and Time (UTC),Receiver,Transmitter,Transmitter Name,Transmitter Serial,Sensor Value,Sensor Unit,Station Name,Latitude,Longitude

use strict;
use warnings;

use Date::Calc qw(Add_Delta_YMDHMS);
use Carp qw(croak); # You could alternatively leave off this line and change all "croak"s to "die" and "carps" to "warn" - but less debugging information is available then

# set up warnings: non-critical error messages to be output at the end of the file process.
my @WARNINGS;
$SIG{__WARN__} = sub { push @WARNINGS, shift };

# with "strict" we need to declare all variables prior to use
my ($line, $junk, $year, $mon, $day, $hh, $mm, $ss, $MMDDYYYY, $HHMMSS, $outdate, $sernumstring, $sernum, $codespacetag, $tech, $spce, $tagcode, $codespace, $data1, $units1, $data2, $units2, $outstring, $datestring, $ver);

# read in header line, look at number of commas to decide what the format is. For another day: handle VUE 1.0 and/or handle VUE 2.0 sans header.
my $coms = (<STDIN> =~ tr/,//);
if ($coms == 13) { $ver = 1; }
elsif ($coms == 9) { $ver = 2; }
else { croak("Unrecognized VUE CSV format. This script written primarily for 'default' format in VUE 2.1.3"); }
# no switch/case in Perl by default?  who knew?

# work on the rest of the lines.
while ($line = <STDIN>) {
  $line =~ s/\r?\n|\r/,,,,,,,,,,,,,,1/g; # instead of cho(m)p for endline processing, this removes any and all endlines of DOS or linux CSVs - the commas and the placeholder at end are to fill out incomplete lines
  my @fields = split(/,/, $line);

  # BLOCK: Process the date and time, same for CSV1 and 2
  $datestring = $fields[0];
  ($year, $mon, $day, $hh, $mm, $ss) = split(/[:\-,\s\/]/, $datestring) or croak("ERROR: unrecognized Date-Time format in $datestring");
  if ($day > 31) { #likely MM/DD/YYYY instead of YYYY-MM-DD
	$junk = $year; $year = $day; $day = $mon; $mon = $junk;
  }
  # Opening in excel and re-saving can trim off the seconds field. You probably wouldn't want this data, but that's a personal call.
  if (!$hh) {$hh="00";} if (!$mm) {$mm="00";} if (!$ss) {$ss="00"; carp("WARNING: date incomplete: $datestring"); }
  # Date-time is UTC, convert to PST 
  ($year, $mon, $day, $hh, $mm, $ss) = Add_Delta_YMDHMS($year, $mon, $day, $hh, $mm, $ss, 0,0,0,-8,0,0) or croak("ERROR $!: $year, $mon, $day, $hh, $mm, $ss"); # consider carp/warn() here
  # Format the date one way MS Access and SQL SERVER like it.
  $MMDDYYYY = join("/", &pad($mon), &pad($day), $year);
  $HHMMSS = join(":", &pad($hh), &pad($mm), &pad($ss));
  $outdate = join(" ", $MMDDYYYY, $HHMMSS);

  # BLOCK: Process the serial number: drop "VR2W" part from CSV2 - we don't need that
  if ($ver == 2) {
	$sernumstring = $fields[1];
	(undef, $sernum) = split(/-/, $sernumstring);
  } else { #ver1
	$sernum = $fields[10];
  }

  # BLOCK: Process the codespace and tag by splitting it apart and then re-joining KHz (A69) with the codespace ID (e.g. 1303)
  if ($ver == 2) {
	$codespacetag = $fields[2];
	($tech, $spce, $tagcode) = split(/-/, $codespacetag);
	$codespace = join("-",$tech,$spce);
  } else {
	$codespace = $fields[1];
	$tagcode = $fields[2];
  }

  # BLOCK: Process data fields for sensor data (pressure, temp, acceler)
  if ($ver == 2) {
	$data1 = $fields[5];
	$units1 = $fields[6];
	# Based on VEMCO e-mail exchange the other sensor should have a different timestamp, I'm not sure why it differed in the era of VUE 1
	$data2 = '';
	$units2 = '';
  } else {
	$data1 = $fields[3];
	$units1 = $fields[4];
	$data2 = $fields[5];
	$units2 = $fields[6];
  }
  # output the line to standard output, which is intercepted by the DOS .BAT file
  $outstring = join(",",($tagcode,$codespace,$outdate,$sernum,$data1,$units1,$data2,$units2));
  print "$outstring\n";
}

# Process any warnings after the file has been read in completely. An error/croak would have terminated the processing of the file immediately.
END { 
  if ( @WARNINGS )  {       
	print STDERR "There were warnings!\n";
	foreach (@WARNINGS) {        
		print STDERR "$_\n";
	}
	exit 1;
  }
}

sub pad {
  sprintf("%02d", $_[0]);
}

