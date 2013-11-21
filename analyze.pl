#!/usr/bin/perl

use strict;
use warnings;
use Data::Dumper;
BEGIN {push @INC, '/Users/kuoj/perl5/lib/perl5'}
use Excel::Writer::XLSX;
use List::Util 'max';

our %patient_name;
our %patient_data;
our %segment_order;

#@first_array = (1, 3, 5, 8, 9);
#@second_array = (0, 2, 4, 6, 8);

#print "$_\n" foreach (&max_peak_and_time(\@first_array, \@second_array));
#&scanfile("HarrisAnn.3.AP3.Strain.000296160-21-EBY7ULOS-results.xls");
foreach ( <*.xls> ) {
	&scanfile($_);
}

#print Dumper(\%patient_data);
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

#print Dumper(\%segment_order);
&write_excel('test-' . time() . '.xlsx');



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
	my $sample_type = $1;
	$build_segment_order = 1 if !exists $segment_order{$sample_type};
	#print $build_segment_order . "\n";
	#print "$file = (AP" . $sample_type . ")\n";

	while (my $line = <XLS>) {
		chomp $line;

		if ( $mode eq "PREAMBLE" ) {
			if ( $line =~ /^Patient Name/ ) {
				$line =~ s/^\s*(.*?)\s*$/$1/;
				$line =~ /^.*?:\s*(\S.*)$/;
				$name = $1;
				#print "name = $name...\n";
			} elsif ( $line =~ /^Patient ID/ ) {
				$line =~ /^.*?:\s*(\d+)/;
				$id = $1;
				#print "id = $id\n";

				# add new id-name mapping
				$patient_name{$id} = $name if !exists $patient_name{$id};

				# create new patient record hash, if needed
				$patient_data{$id} = {} if !exists $patient_data{$id};

				$rec_ref = {};
				$patient_data{$id}->{$sample_type} = $rec_ref;
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
				if ( $line =~ /^Global/ ) {
					#print "Summary starts at line $." . "\n";
					$mode = "SUMMARY";
					$build_segment_order = 0 if $build_segment_order;
				} else {
					#print "Segment '" . $line . "' at line $.\n";
					$seg = $line;
					$time_array_ref = [];
					$data_array_ref = [];
					push @{$segment_order{$sample_type}}, $seg if $build_segment_order;
					$mode = "HEADING";
				}
			}
		} elsif ( $mode eq "HEADING" ) {
			$heading_ref = [ split '\t', $line ];
			$mode = "DATA";
		} elsif ( $mode eq "DATA" ) {
			if ( $line !~ /^$/ ) {
				if ( $line =~ /^--------/ ) {
					#print "Divider at line $." . "\n";
					$rec_ref->{$seg} = [$heading_ref, $time_array_ref, $data_array_ref];
					$mode = "SEGMENT";
				} else {
					my ($t, $v) = split '\t', $line;
					#print "\tPushing time $t, value $v...\n" if $seg eq 'BAS Long. Strain';
					push @{$time_array_ref}, $t;
					push @{$data_array_ref}, $v;
				}
			}
		}
	}

	close XLS;
}




# given two references to arrays for time and stress, return peak stress and time to peak
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

	# define cell formats
	my $ap_format = $workbook->add_format(bold => 1, align => 'center', left => 1);
	my $div_format = $workbook->add_format(left => 1);
	my $bottom_format = $workbook->add_format(bottom => 1);
	my $corner_format = $workbook->add_format(left => 1, bottom => 1);

	# create headers by iterating across chambers and their segments
	my $segment_x_start = 2, my $segment_x_end;
	my $chamber_x_start = $segment_x_start, my $chamber_x_end;
	my $seg_count = 0;
	foreach my $chamber (sort keys %segment_order) {
		my @filtered_segments = grep $_ =~ /Strain$/, @{$segment_order{$chamber}};
		$chamber_x_end = $chamber_x_start - 1 + 2*(scalar @filtered_segments);

		$worksheet->set_column($chamber_x_start, 0, undef, $div_format);
		$worksheet->merge_range(0, $chamber_x_start, 0, $chamber_x_end, $chamber, $ap_format);

		foreach my $segment (@filtered_segments) {
			$segment_x_end = $segment_x_start + 1;

			# segment header formatting logic
			my $new_seg_format = $workbook->add_format(align => 'center', right => 4);
			$new_seg_format->set_left() if $segment_x_start==$chamber_x_start;
			$new_seg_format->set_bg_color($seg_count%2==0 ? 22 : 41); # alternating light gray background

			# write segment name
			$worksheet->merge_range(1, $segment_x_start, 1, $segment_x_end, $segment, $new_seg_format);

			$worksheet->write(2, $segment_x_start, 'peak', ($segment_x_start==$chamber_x_start ? $corner_format : $bottom_format));
			# 'time' header cell should have solid bottom border, dotted right border 
			$new_seg_format = $workbook->add_format(bottom => 1);
			$new_seg_format->set_right(4);
			$worksheet->write(2, $segment_x_start+1, 'time', $new_seg_format);

			# increment positions and counters
			$segment_x_start = $segment_x_end + 1;
			$seg_count++;
		}

		$chamber_x_start = $chamber_x_end + 1;
	}
	$worksheet->write(2, 0, 'Patient ID', $bottom_format);
	$worksheet->write(2, 1, 'Name', $bottom_format);

	# resize name column
	my $longest_name_size = max map length, values %patient_name;
	$worksheet->set_column(1, 0, $longest_name_size);

	# write data
	my $data_x_start; # stating x-position (column)
	my $data_y = 3; # running y-position (row)
	foreach my $patient_id (keys %patient_data) {
		$worksheet->write_string($data_y, 0, $patient_id); # write ID as string to preserve leading zeros
		$worksheet->write_string($data_y, 1, $patient_name{$patient_id});
		$data_x_start = 2;
		foreach my $chamber (sort keys %segment_order) {
			foreach my $segment (grep $_ =~ /Strain$/, @{$segment_order{$chamber}}) {
				my $segment_ref = $patient_data{$patient_id}->{$chamber}->{$segment};
				if ( $segment_ref ) {
					my ($peak, $t) = &max_peak_and_time($segment_ref->[1], $segment_ref->[2]);
					$worksheet->write($data_y, $data_x_start++, $peak);
					$worksheet->write($data_y, $data_x_start++, $t, $workbook->add_format(right => 4));
				} else {
					$worksheet->write($data_y, ++$data_x_start, undef, $workbook->add_format(right => 4));
					$data_x_start++;
				}
			}
		}
		$data_y++;
	}

}
