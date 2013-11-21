#!/usr/bin/perl

use strict;
use warnings;
use Data::Dumper;
use Getopt::Long;
BEGIN {push @INC, '/Users/kuoj/perl5/lib/perl5'}
use Excel::Writer::XLSX;
use List::Util 'max';

our %patient_name;
our %patient_data;
our %segment_order;

# workaround for printing Unicode output
binmode STDOUT, ":encoding(UTF-16LE)";

#@first_array = (1, 3, 5, 8, 9);
#@second_array = (0, 2, 4, 6, 8);

#print "$_\n" foreach (&max_peak_and_time(\@first_array, \@second_array));
#&scanfile("HarrisAnn.3.AP3.Strain.000296160-21-EBY7ULOS-results.xls");

# command-line options
my $debug_state_machine = 0;
my $debug_parser = 0;
my $verbose = 0;
my $print_results = 0;
my $create_excel = 1;
GetOptions(
	'debug-state' => \$debug_state_machine,
	'debug-parser' => \$debug_parser,
	'verbose' => \$verbose,
	'print-results' => \$print_results,
	'excel!' => \$create_excel);



foreach ( <*.xls> ) {
	&scanfile($_);
}

#print Dumper(\%patient_data);
if ( $print_results == 1 ) {
	foreach my $id (keys %patient_data) {
		print "Patient $id ($patient_name{$id})\n";
		foreach my $chamber (sort keys %{$patient_data{$id}}) {
			print "\t$chamber\n";
			foreach my $segment (keys %{$patient_data{$id}->{$chamber}}) {
				next if $segment !~ /Strain$/;
				print "\t\t$segment";
				my $segment_ref = $patient_data{$id}->{$chamber}->{$segment};
				print " (" . scalar @{$segment_ref->[1]} . ") ";
				my ($peak, $t) = &max_peak_and_time($segment_ref->[1], $segment_ref->[2]);
				print "\t\tmax $peak\% @ $t s.\n";
			}
		}
	}
}

#print Dumper(\%segment_order);
if ( $create_excel == 1 ) {
	&write_excel('test-' . time() . '.xlsx');
}


sub scanfile {
	my $file = shift;
	my $mode = "PREAMBLE";
	my ($id, $name);
	my $build_segment_order = 0;
	my $seg;
	my $rec_ref;
	my ($heading_ref, $time_array_ref, $data_array_ref);

	open XLS, "<:encoding(UTF-16LE)", $file || die "Cannot open $file\n";

	# determine AP2, AP3, AP4, SAXM, or SAXB
	$file =~ /.*?\.(AP\d|SAX\w)\..*/;
	my $view_type = $1;
	if ( not defined $view_type ) {
		print "*** Skipping $file; cannot determine view (AP2, AP3, etc.) from filename\n";
		return;
	}
	$build_segment_order = 1 if !exists $segment_order{$view_type};
	if ( $debug_parser == 1 ) {
		#print $build_segment_order . "\n";
		print "===== $file = (" . $view_type . ") =====\n";
	}

	while (my $line = <XLS>) {
		chomp $line;

		if ( $mode eq "PREAMBLE" ) {
			if ( $line =~ /^Patient Name/ ) {
				$line =~ s/^\s*(.*?)\s*$/$1/;
				$line =~ /^.*?:\s*(\S.*)$/;
				$name = $1;
				print "\tname = $name...\n" if $debug_parser == 1;
			} elsif ( $line =~ /^Patient ID/ ) {
				$line =~ /^.*?:\s*(\d+)/;
				$id = $1;
				print "\tid = $id\n" if $debug_parser == 1;

				# add new id-name mapping
				$patient_name{$id} = $name if !exists $patient_name{$id};

				# create new patient record hash, if needed
				$patient_data{$id} = {} if !exists $patient_data{$id};

				$rec_ref = {};
				$patient_data{$id}->{$view_type} = $rec_ref;
			} elsif ( $line =~ /^--------/ ) {
				$mode = "FIRST_DIVIDER";
			}
		} elsif ( $mode eq "FIRST_DIVIDER" ) {
			if ( $line =~ /^$/ ) {
				# empty line (gap)
				$mode = "SEGMENT";
			}
		} elsif ( $mode eq "SEGMENT" ) {
			if ( $line !~ /^$/ ) {
				if ( $line =~ /\s=\s/ ) {
					print "\t* Summary starts at line $." . "\n" if $debug_parser == 1;
					$mode = "SUMMARY";
					$build_segment_order = 0 if $build_segment_order;
				} else {
					print "\t* Segment '" . $line . "' at line $.\n" if $debug_parser == 1;
					$seg = $line;
					$time_array_ref = [];
					$data_array_ref = [];
					push @{$segment_order{$view_type}}, $seg if $build_segment_order;
					$mode = "HEADING";
				}
			}
		} elsif ( $mode eq "HEADING" ) {
			$heading_ref = [ split '\t', $line ];
			$mode = "DATA";
		} elsif ( $mode eq "DATA" ) {
			if ( $line !~ /^$/ ) {
				if ( $line =~ /^--------/ ) {
					print "\t\t(Divider at line $." . ")\n" if $debug_parser == 1;
					$rec_ref->{$seg} = [$heading_ref, $time_array_ref, $data_array_ref];
					$mode = "SEGMENT";
				} else {
					my ($t, $v) = split '\t', $line;
					print "\t\tPushing time $t, value $v...\n" if $debug_parser == 1 and $verbose == 1;
					push @{$time_array_ref}, $t;
					push @{$data_array_ref}, $v;
				}
			}
		}

		if ( $debug_state_machine == 1 ) {
			printf "%-15s\t%s\n", '['.$mode.']', $line if $line !~ /^$/;
		}
	}

	close XLS;
}




# given two references to arrays for time and stress, return peak stress and time to peak
# (peak is actually negative max!)
sub max_peak_and_time {
	my ($t, $s) = @_; 
	my $max_peak = $s->[0], my $time_to_peak = $t->[0];

	for (my $i = 0; $i < @$s; $i++) {
		if ( $s->[$i] < $max_peak ) {
			$max_peak = $s->[$i];
			$time_to_peak = $t->[$i];
		}
	}

	return ($max_peak, $time_to_peak);
}




# write Excel file (XLSX format)
sub write_excel {
	my $filename = shift;

	# create workbook and add worksheet
	my $workbook = Excel::Writer::XLSX->new($filename);
	my $worksheet = $workbook->add_worksheet();
	$worksheet->freeze_panes(3, 2);

	# define cell formats
	my $view_format = $workbook->add_format(bold => 1, align => 'center', left => 1);
	my $div_format = $workbook->add_format(left => 1);
	my $bottom_format = $workbook->add_format(bottom => 1);
	my $corner_format = $workbook->add_format(left => 1, bottom => 1);

	# print patient header and info
	$worksheet->write(2, 0, 'Patient ID', $bottom_format);
	$worksheet->write(2, 1, 'Name', $bottom_format);
	my $current_row = 3; # running y-position (row)
	foreach my $patient_id (keys %patient_data) {
		#$worksheet->write_string($current_row, 0, $patient_id); # write ID as string to preserve leading zeros
		$worksheet->write($current_row, 0, $patient_id, $workbook->add_format(num_format => '000000000'));
		$worksheet->write_string($current_row, 1, $patient_name{$patient_id});
		$current_row++;
	}
	# resize name column
	my $longest_name_size = max map length, values %patient_name;
	$worksheet->set_column(1, 0, $longest_name_size);

	# iterate across views, print segment headers and data for each patient
	my $segment_x_start = 2, my $segment_x_end;
	my $view_x_start = $segment_x_start, my $view_x_end;
	my $seg_count = 0;
	foreach my $view (sort keys %segment_order) {
		my @filtered_segments = grep $_ =~ /Strain$/, @{$segment_order{$view}};
		$view_x_end = $view_x_start - 1 + 2*(scalar @filtered_segments);

		$worksheet->set_column($view_x_start, 0, undef, $div_format);
		$worksheet->merge_range(0, $view_x_start, 0, $view_x_end, $view, $view_format);

		foreach my $segment (@filtered_segments) {
			$segment_x_end = $segment_x_start + 1;

			# segment header formatting logic
			my $new_seg_format = $workbook->add_format(align => 'center', right => 4);
			$new_seg_format->set_left() if $segment_x_start==$view_x_start;
			$new_seg_format->set_bg_color($seg_count%2==0 ? 22 : 41); # alternating light gray background

			# write segment name
			$worksheet->merge_range(1, $segment_x_start, 1, $segment_x_end, $segment, $new_seg_format);

			$worksheet->write(2, $segment_x_start, 'peak', ($segment_x_start==$view_x_start ? $corner_format : $bottom_format));
			# 'time' header cell should have solid bottom border, dotted right border 
			$new_seg_format = $workbook->add_format(bottom => 1);
			$new_seg_format->set_right(4);
			$worksheet->write(2, $segment_x_start+1, 'time', $new_seg_format);

			# iterate through patients for current segment data
			my $current_row = 3; # running y-position (row)
			foreach my $patient_id (keys %patient_data) {
				my $segment_ref = $patient_data{$patient_id}->{$view}->{$segment};
				if ( $segment_ref ) {
					my ($peak, $t) = &max_peak_and_time($segment_ref->[1], $segment_ref->[2]);
					$worksheet->write($current_row, $segment_x_start, $peak);
					$worksheet->write($current_row, $segment_x_start+1, $t, $workbook->add_format(right => 4));
				} else {
					$worksheet->write($current_row, $segment_x_start+1, undef, $workbook->add_format(right => 4));
				}
				$current_row++;
			}

			# increment positions and counters
			$segment_x_start = $segment_x_end + 1;
			$seg_count++;
		}

		$view_x_start = $view_x_end + 1;
	}
}
