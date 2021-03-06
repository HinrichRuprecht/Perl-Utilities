#!/usr/bin/perl -w

use strict;

my $description=<<'DESC';
Extract sheet(s) from a Libre/OpenOffice spreadsheet file to csv format.

NAME
    ExtractODS.pl (source)

SYNOPSIS
    perl ExtractODS.pl [Options] SpreadsheetFile [SheetName ...]

DESCRIPTION
    The program reads the spreadsheet contents of the given file and
    writes only the cells values to STDOUT (or the specified output file).
    If sheet names are specified, only these sheets are extracted.
    Otherwise, only the first sheet or, when the output filename contains
    <SHEET> as placeholder for the sheet name, all sheets are extracted.
    SheetName may contain wildcards as in regular expressions, e.g.
    "Project.*" will extract all sheets with names starting with 'Project'.
    
    Only Libre/OpenOffice spreadsheet files are supported (.ods/.ots/.fods),
    i.e. files which are (internally) stored in one main XML file.
    Other file formats which use XML (e.g. .xlsx) separate text and structure.
    To extract sheets from Excel file formats, they must first be converted
    to .ods, e.g. with soffice --convert-to ods --headless --norestore ABC.xls
    If Libre/OpenOffice is found, the conversion is also done (with temporary
    files).

OPTIONS
    -h      : Show this help information
    -v      : Verbose
    -f      : Convert non-Libre/OpenOffice file to .ods without asking
              (only if soffice is found)
    -o FILE : Output file (default STDOUT)
              The filename may contain <SHEET> as placeholder for the
              sheet name.
    -S CHAR : Separator (Default horizontal tab)
    -D CHAR : Delimiter (Default '"'). Used when a value contains Separator.

AUTHOR
    © 2018 Hinrich Ruprecht : usable according to GNU public license
    
DESC
my $version="1.0.1";

# Read Options:
my $usage="Usage: $^X $0 [-hv] \n\t[-o OUTFILE] \n\t[-S 'SEPARATOR'] "
    ."[-D 'DELIMITER'] \n\tODSFILE [SHEET ...]";
my $verbose=0; my $test=0; my $target=""; my $sep="\t"; my $del='"'; 
my $force=0;
if ($#ARGV<0 || ($#ARGV==0 && $ARGV[0] eq "")) {
    print $usage,"\nArguments (-h or ? for help): ";
    my $tmp=<STDIN>;
    chomp($tmp) if $tmp;
    exit unless $tmp;
    @ARGV=split(" ",$tmp);
    }
my $par;
while ($#ARGV>=0 && (($par=shift) eq "" || substr($par,0,1) eq "-")) {
    next unless $par;
    $verbose++ if $par=~/v/;
    $test=1 if $par=~/t/;
    $force=1 if $par=~/f/;
    if ($par=~/h/) { print $description,"\n"; exit; }
    if ($par=~/([oSD])/) {
        my $par=$1; my $val=shift;
        $val=$2 if $val=~/^([\"\'])(.*)([\"\'])$/ && $1 eq $3;
        if ($par eq "o") { $target=$val; }
        elsif ($par eq "S") { $sep=$val; }
        elsif ($par eq "D") { $del=$val; }
        }
    }
print STDERR "par=$par\n" if $test>0;
if ($par eq "?" || $par eq "0") { print $description,"\n"; exit; }
$verbose++ if $test>0;
# Parameter(s):
my $spreadsheetFile = $par;

die "No spreadsheet file specified\nUse -h for help\n" unless $spreadsheetFile;
die "File $spreadsheetFile not found\n" if !-e $spreadsheetFile;
die "No read access to $spreadsheetFile\n" if !-r $spreadsheetFile;

print STDERR $0," version ",$version,"\n" if $verbose>0;
die "Delimiter and Separator must be different\n" if $del eq $sep;

my $tmpName="tmp";

extractSheets($spreadsheetFile,$target,\@ARGV);

sub extractSheets {
    my ($spreadsheetFile,$output,$aSheets) = @_;
    
    my $zipTypes="ods|ott"; # xOffice files use zip format to store contents
    my $xmlTypes="fods"; # flat ODF XML spreadsheet

    # replacements for &...;
    my %repl; 
        $repl{"amp"}='&'; $repl{"lt"}='<'; $repl{"gt"}='>';
        $repl{"apos"}="'"; $repl{"quot"}='"';
    
    my $XMLfile;
    if ($spreadsheetFile=~/\.($xmlTypes)$/i) # fods/fots? : 
        { $XMLfile=$spreadsheetFile; }
    else {
        if ($spreadsheetFile!~/\.($zipTypes)$/i) {
            # convert2ods exits if soffice exists but an error occurred.
            $spreadsheetFile=convert2ods($spreadsheetFile); 
            }
        # Unzip contents from spreadsheetFile
        use Archive::Zip qw( :ERROR_CODES :CONSTANTS );
        my $zip = Archive::Zip->new();
        if ($zip->read( $spreadsheetFile ) != AZ_OK) {
            print STDERR "File $spreadsheetFile probably has wrong format.\n",
                "Only Libre/OpenOffice spreadsheet files are supported. ",
                "Use Libre/OpenOffice to convert to .ods!\n";
            exit;
            }
        if (-e "$tmpName.xml") { 
            unlink("$tmpName.xml") 
            || die "Can't remove $tmpName.xml (used as temporary file)\n$!\n";
            }
        $zip->extractMember("content.xml","$tmpName.xml");
        if (!-e "$tmpName.xml") 
            { die "Can't extract content.xml from $spreadsheetFile\n"; }
        $XMLfile="$tmpName.xml";
        }
    
    # Check for XML header and read contents (normallly in 2nd line):
    open(XML,$XMLfile) 
        || die "Can't open contents of $spreadsheetFile\n$!\n"; 
    my $xmlHeader=<XML>;
    if ($xmlHeader!~/\<\?xml version/) {
        print STDERR "* Header of $spreadsheetFile content is not XML\n";
        }
    my @content=<XML>; # should only be one line
    close(XML);
    my $content=join("",@content);
    my $p1=index($content,"<office:spreadsheet");
    die "No spreadsheet in $spreadsheetFile\n" if $p1<0;
    
    # Escape special characters in sheet name parameters:
    for (my $iS=0; $iS<=$#$aSheets; $iS++) {
        $aSheets->[$iS]=~s/([\-\+])/\\$1/g;
        }
    
    # Just look for 
    #   table:name  -> new sheet 
    #   table:row   -> new line, i.e. write cells from previous row
    #   table:cell  -> only take text between text:p tags
    #   ..repeated  -> remember cell or row count for next output
    #                   (don't write empty cells at end of line, or empty
    #                   rows at end of sheet)
    my $nCells=0; my $line=""; my $takeSheet=0; my $nSheets=0; my $nRows=0;
    while (($p1=index($content,"<",$p1))>0) {
        my $p2=index($content,">",$p1);
        my $tag=substr($content,$p1+1,$p2-$p1-1);
        $p1=$p2+1;
        if ($tag=~/table\:name\=\"([^\"]+)\"/) {
            my $sheetName=$1;
            print STDERR "* sheet '$sheetName' " if $test>0;
            last if $nSheets>0 && $output!~/\<SHEET\>/ 
                 && ($#$aSheets<0 || $aSheets->[0] eq "");
            $takeSheet=takeSheet($sheetName,$output,$aSheets);
            if ($takeSheet>0) { $nSheets++; $nRows=0; }
            next;
            }
        elsif ($takeSheet==0) { next; }
        if (substr($tag,0,1) eq "/" || substr($tag,-1) eq "/") {
            if ($tag=~/\:table-row/) {
                if ($line ne "") {
                    foreach (1..$nRows) { print "\n"; }
                    print $line,"\n";
                    $nRows=0;
                    $line="";
                    }
                else { $nRows++; }
                $nCells=0; 
                $nRows+=($tag=~/repeated\=\"(\d+)\"/ ? $1-1 : 0);
                }
            elsif ($tag=~/\:table-cell/) {
                $nCells+=($tag=~/repeated\=\"(\d+)\"/ ? $1 : 1);
                }
           }
        elsif ($tag=~/\:table-row/) {
            $nRows+=($tag=~/repeated\=\"(\d+)\"/ ? $1-1 : 0);
            }
        elsif ($tag eq "text:p") {
            for (my $i=0; $i<$nCells; $i++) { $line.=$sep; }
            $nCells=0; my $p3; my $p4=$p2;
            while (($p3=index($content,"</text:p>",$p4))>0
                    && ($p4=index($content,"<",$p3+1))>0 
                    && substr($content,$p4+1,6) eq "text:p") 
                { } #find last </text:p> in cell
            # Text between text:p tags may contain further tags -> remove:
            my $val=substr($content,$p2+1,$p3-$p2-1);
            while ($val=~/^(.*)\<[^\>]+\>(.*)$/) { $val=$1.$2; }
            # Replace &...; by their replacement character
            while ($val=~/^(.*)\&([a-z]+)\;(.*)$/) {
                my $s=$2; my $r=$repl{$s};
                if (!defined($r)) {
                    $r=uc($s);
                    print STDERR "* No replacement for &$s\; in $val\n";
                    }
                $val=$1.$r.$3;
                }
            $p4=0; my $useDel=0;
            # Escape delimiter characters within text (with ")
            while (($p4=index($val,$del,$p4))>=0) {
                $val=substr($val,0,$p4).'"'.substr($val,$p4);
                $p4+=2;
                $useDel=1;
                }
            $val=$del.$val.$del if $useDel>0 || index($val,$sep)>=0;
            $line.=$val;
            $p1=$p3+2;
            }
        }
    close(OUT);
    
    sub convert2ods {
        my ($file) = @_;
        
        use File::Copy;

        my $os=$^O; my $soffice;
        if ($os=~/MSWin/i) # Windows
            { $soffice=glob('C:\Program\ Files*\*Office*\*\soffice.exe'); }
        elsif ($os=~/x$|bsd/) # Unix/Linux 
            { $soffice="/usr/bin/soffice"; }
        else { $soffice="soffice"; }
        return $file if $soffice ne "soffice" && !-e $soffice;
        my $type=($file=~/\.([\w]+)$/ ? $1 : "");
        copy($file,"$tmpName.$type") 
            || die "Can't copy $file to $tmpName.$type\n$!\n";
        $soffice='"'.$soffice.'"' if index($soffice,'"')>=0;
        my $cmd=$soffice." --convert-to ods $tmpName.$type "
                ."--headless --norestore";
        if ($verbose>0 || $force==0) {
            print STDERR "File type .$type is not supported.\n",
                "Convert $file to .ods using (on a copy)\n  ",$cmd,"\n",
                "Libre/OpenOffice must not be running! ";
            if ($force==0) {
                print STDERR "(Y/n): ";
                my $tmp=<STDIN>;
                exit if !defined($tmp) || $tmp=~/^[nq]/i;
                }
            else { print STDERR "\n"; }
            }
        if (-e "$tmpName.ods") {
            unlink("$tmpName.ods") 
                || die "Can't remove temporary file '$tmpName.ods\n$!\n";
            }
        system($cmd);
        die "$tmpName.ods not created\n" if !-e "$tmpName.ods";
        return "$tmpName.ods";
        } # convert2ods

    sub takeSheet { 
    # Check wether sheet is wanted, and open output file
        my ($sheetName,$output,$aSheets) = @_;
        
        my $res=0;
        for (my $iS=0; $iS<=$#$aSheets; $iS++) {
            my $sheet_=$aSheets->[$iS];
            if ($sheetName=~/^$sheet_$/ || $sheetName eq $sheet_) {
                print STDERR "* $sheetName matches $sheet_\n" if $test>0;
                $res=1;
                last;
                }
            }
        if ($res==0 && $sheetName!~/(__Anonymous|BuiltIn__|_xlfn_ISFORMULA)/
            && ($#$aSheets<0 || $aSheets->[0] eq "")) 
            { $res=1; }
        if ($res>0) {
            select STDOUT;
            $output=~s/\<SHEET\>/$sheetName/g;
            if ($verbose>0) {
                print "* Extracting sheet $sheetName",
                      ($output ne ""? " to ".$output : ""),
                      "\n";
                }
            if ($output ne "") {
                close(OUT);
                open(OUT,">".$output) || die "Can't write to $output\n$!\n";
                select OUT;
                }
            }
        return $res;
        } # takeSheet
        
    } # extractSheets
