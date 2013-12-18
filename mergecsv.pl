#<!-- -*- encoding: utf-8n -*- -->
use strict;
use warnings;
use utf8;

use Encode;
use Encode::JP;

use Data::Dumper;
use Cwd;

use Spreadsheet::ParseExcel;
use Spreadsheet::WriteExcel;

use File::Basename;

use Tk;
use Tk::Tree;
use Tk::FileSelect;
use Tk::NoteBook;
use Tk::ProgressBar;
use Win32::Process;
use Win32::GUI();

# ファイル入出力を制御する
use open IN  => ":encoding(cp932)";
use open OUT => ":encoding(cp932)";
# 標準入出力を制御する
binmode STDIN, ':encoding(cp932)';
binmode STDOUT, ':encoding(cp932)';
binmode STDERR, ':encoding(cp932)';

my $gCSVPattern = '[^\",]*|\"(?:[^\"]|\"\")*\"';

my $version = "mergecsv ver. 0.13.08.24.\n";
print "$version\n";
my ($argv, $gOptions) = getOptions(\@ARGV); # オプションを抜き出す

my %gTk = ();

if( !exists $gOptions->{'nogui'} ){
	# gui mode.
	my $font=['MS　ゴシック', 12, 'normal'];
	
	$gTk{'window'} = MainWindow->new();
	$gTk{'window'}->title("$version");
#	$gTk{'window'}->geometry("700x440");

	%{$gTk{'work'}} = ();
	my $work = \%{$gTk{'work'}};
	my $cwd = decode('cp932', Cwd::getcwd());
	$cwd =~ s/\//\\/g;

	$work->{'out_path'} = '';
	$work->{'pattern_file'} = '';
	$work->{'search_directory'} = $cwd;
	$work->{'expected_path'} = $cwd;
	$work->{'expected_xls_path'} = '';
	$work->{'use_xls_expected'} = 0;
	if( -f ".mergecsv.ini" ){
		loadINI(".mergecsv.ini");
	}

	$gTk{'menubar'} = $gTk{'window'}->Menu( -type => 'menubar' );
	$gTk{'window'}->configure( -menu => $gTk{'menubar'} );
	
	$gTk{'menu_file'} = $gTk{'menubar'}->cascade(-label => 'file', -under => 0, -tearoff => 0);
#	$gTk{'menu_help'} = $gTk{'menubar'}->cascade(-label => 'help', -under => 0, -tearoff => 0);
	
	# file list.
	$gTk{'menu_file'}->command(-label => 'load', -under => 0, -command => \&eventLoadINI );
	$gTk{'menu_file'}->command(-label => 'save', -under => 0, -command => \&eventSaveINI );
	$gTk{'menu_file'}->separator;
	$gTk{'menu_file'}->command(-label => 'exit', -under => 0, -command => \&exit );
	
	# help list.
#	$gTk{'menu_help'}->command(-label => 'version', -under => 0, -command => \&Version );
	
	
	# マージフレーム
	$gTk{'merge_frame'} = $gTk{'window'}->Frame(-relief => 'flat', -borderwidth => 1)->pack( -fill => 'x',  -padx => 5, -pady => 5);
	
	$gTk{'out_path_frame'} = $gTk{'merge_frame'}->Frame(-relief => 'flat', -borderwidth => 0)->pack( -fill => 'x', -padx => 2, -pady => 2);
	$gTk{'out_path_frame'}->Button( -font=>$font, -width => 15, -height => 1, -text => "出力ファイル名", -command => [ \&fileSelector, \$work->{'out_path'}, 'xls', 'save' ])->pack(-side => 'right', -padx => 2,  -pady => 2);
	$gTk{'out_path_frame'}->Entry( -font=>$font,  -textvariable => \$work->{'out_path'} )->pack( -fill => 'x', -padx => 2,  -pady => 2);
	
	$gTk{'pattern_file_frame'} = $gTk{'merge_frame'}->Frame(-relief => 'flat', -borderwidth => 0)->pack( -fill => 'x', -padx => 2, -pady => 2);
	$gTk{'pattern_file_frame'}->Button( -font=>$font, -width => 15, -height => 1, -text => "取得パターン", -command => [ \&fileSelector, \$work->{'pattern_file'}, 'csv','open', 1 ])->pack(-side => 'right', -padx => 2,  -pady => 2);
	$gTk{'pattern_file_frame'}->Entry( -font=>$font,  -textvariable => \$work->{'pattern_file'} )->pack( -fill => 'x', -padx => 2,  -pady => 2);
	
	$gTk{'search_directory_frame'} = $gTk{'merge_frame'}->Frame(-relief => 'flat', -borderwidth => 0)->pack( -fill => 'x', -padx => 2, -pady => 2);
	$gTk{'search_directory_frame'}->Button( -font=>$font, -width => 15, -height => 1, -text => "検索ディレクトリ", -command => [ \&pathSelector, \$work->{'search_directory'} ])->pack(-side => 'right');
	$gTk{'search_directory_frame'}->Entry( -font=>$font,  -textvariable => \$work->{'search_directory'} )->pack( -fill => 'x', -padx => 2,  -pady => 2);

	$gTk{'merge_frame'}->Button( -font=>$font, -width => 15, -height => 1, -text => "収集", -command => [ \&eventGetActual ])->pack(-side => 'right', -padx => 2, -pady => 2);

	# 突合せフレーム
	$gTk{'expected_frame'} = $gTk{'window'}->Frame(-relief => 'flat', -borderwidth => 1)->pack( -fill => 'x', -padx => 5, -pady => 5);
	
	$gTk{'expected_path_frame'} = $gTk{'expected_frame'}->Frame(-relief => 'flat', -borderwidth => 0)->pack( -fill => 'x', -padx => 2, -pady => 2);
	$gTk{'expected_path_frame'}->Button( -font=>$font, -width => 15, -height => 1, -text => "予想結果ディレクトリ", -command => [ \&pathSelector, \$work->{'expected_path'} ])->pack(-side => 'right');
	$gTk{'expected_path_frame'}->Entry( -font=>$font,  -textvariable => \$work->{'expected_path'} )->pack( -fill => 'x', -padx => 2,  -pady => 2);

	$gTk{'xls_file_frame'} = $gTk{'expected_frame'}->Frame(-relief => 'flat', -borderwidth => 0)->pack( -fill => 'x', -padx => 2, -pady => 2);
	$gTk{'xls_file_frame'}->Button( -font=>$font, -width => 15, -height => 1, -text => "予想結果xlsファイル", -command => [ \&fileSelector, \$work->{'expected_xls_path'}, 'xls', 'open' ])->pack(-side => 'right');
	$gTk{'xls_file_frame'}->Entry( -font=>$font,  -textvariable => \$work->{'expected_xls_path'} )->pack( -fill => 'x', -padx => 2,  -pady => 2);

	$gTk{'expected_frame'}->Button( -font=>$font, -width => 15, -text => "突合", -command => [ \&eventCheck ])->pack(-side => 'right', -padx => 2, -pady => 2);
	$gTk{'expected_frame'}->Checkbutton( -font=>$font, -text => "xlsファイルを使う", -variable => \$work->{'use_xls_expected'} )->pack( -side => 'right', -padx => 2,  -pady => 2);

	# ログフレーム
	$gTk{'log_frame'} = $gTk{'window'}->Frame(-relief => 'flat', -borderwidth => 1)->pack( -fill => 'both', -padx => 2, -pady => 2);
	$gTk{'log'} = $gTk{'log_frame'}->Scrolled('Listbox',
											  -scrollbars=> 'osoe',
											  -background=> 'white',
											  -selectforeground=> 'brown',
											  -selectbackground=> 'cyan',
											  -selectmode=> 'browse');
	$gTk{'log'}->pack( -fill => 'both' );

	# プログレスフレーム
	$gTk{'progress_frame'} = $gTk{'window'}->Frame(-relief => 'flat', -borderwidth => 1)->pack( -fill => 'x', -padx => 2, -pady => 2);
	$gTk{'progress'} = $gTk{'progress_frame'}->ProgressBar(
		-from => 0,
		-to => 100,
		-colors=>[ 0, '#104E8B' ]
		);
	$gTk{'progress'}->pack( -fill => 'x');
	
	MainLoop();
	exit;
}

if( 2 > scalar(@{$argv}) ){
	print "Usage: mergecsv [options] <pattern file name> <search directory>\n";
	print "  options : -expected_path     ... expected directory. need same hierarchy of search directory\n";
	print "          : -expected_xls_path ... expected xls file. need sheet name is \"expected\"\n";
	print "          : -o                 ... output file.\n";
	print "          : -nogui             ... no gui mode.\n";
	print "          : -dbg               ... debug mode\n";
	print "https://github.com/oya3/mergecsv\n";
	exit;
}
mainProcess($argv);
exit;

sub clearProgress
{
	$gTk{'progress'}->value(0);
	$gTk{'window'}->update;
}

sub setProgress
{
	$gTk{'progress'}->value(shift);
	$gTk{'window'}->update;
}

sub eventLoadINI
{
	my $file = '';
	fileSelector( \$file, 'ini', 'open');
	if( $file eq '' ){
		return;
	}
	loadINI($file);
}

sub eventSaveINI
{
	my $file = '';
	fileSelector( \$file, 'ini', 'save');
	if( $file eq '' ){
		return;
	}
	saveINI($file);
}

sub loadINI
{
	my ($file) = @_;
	my $file_sjis = encode('cp932', $file);
	open(IN ,"<$file_sjis");
	my @array = <IN> ;
	close IN;
	my $work = \%{$gTk{'work'}};
	
	foreach my $line (@array){
		if( $line =~ /^(.+?),(.+?)$/ ){
			$work->{$1} = $2;
		}
	}
}

sub saveINI
{
	my $file = shift;
	my $file_sjis = encode('cp932', $file);
	my @items = ('out_path', 'pattern_file', 'search_directory', 'expected_path', 'expected_xls_path','use_xls_expected');
	my $work = \%{$gTk{'work'}};
	my @out = ();
	foreach my $item(@items){
		if( exists $work->{$item} ){
			push @out, "$item,$work->{$item}\n";
		}
	}
 	open (OUT,">$file_sjis");
	foreach my $line (@out){
		print OUT $line;
	}
 	close (OUT);
}

sub eventGetActual
{
	my $work = \%{$gTk{'work'}};
	
	if( isNull( $work, 'pattern_file', '取得パターンを設定しやがれ') ){
		return;
	}
	if( isNull( $work, 'search_directory', '検索フォルダを設定しやがれ') ){
		return;
	}
	if( isNull( $work, 'out_path', '出力ファイルを設定しやがれ') ){
		return;
	}
	
	my @argv = ();
	push @argv, "$work->{'pattern_file'}";
	push @argv, "$work->{'search_directory'}";
	$gOptions->{'o'} = $work->{'out_path'};

	eval{
		clearProgress();
		mainProcess(\@argv);
	};
	if( $@ ){
		$gTk{'window'}->messageBox(
			-type => "ok",
			-icon => "info",
			-title => "mergecsv",
			-message => $@
			);
		#print $@."\n";
	}
	%{$gOptions} = ();
}

sub eventCheck
{
	my $work = \%{$gTk{'work'}};

	if( isNull( $work, 'pattern_file', '取得パターンを設定しやがれ') ){
		return;
	}
	if( isNull( $work, 'search_directory', '検索フォルダを設定しやがれ') ){
		return;
	}
	if( isNull( $work, 'out_path', '出力ファイルを設定しやがれ') ){
		return;
	}
	my $targetExpected = 'expected_path';
	if( $work->{'use_xls_expected'} ){
		$targetExpected = 'expected_xls_path';
	}
	
	if( isNull( $work, $targetExpected, '予想結果フォルダを設定しやがれ') ){
		return;
	}
	
	my @argv = ();
	push @argv, "$work->{'pattern_file'}";
	push @argv, "$work->{'search_directory'}";
	$gOptions->{'o'} = $work->{'out_path'};
	$gOptions->{$targetExpected} = $work->{$targetExpected};
	
	eval{
		clearProgress();
		mainProcess(\@argv);
	};
	if( $@ ){
		$gTk{'window'}->messageBox(
			-type => "ok",
			-icon => "info",
			-title => "mergecsv",
			-message => $@
			);
		#print $@."\n";
	}
	%{$gOptions} = ();
}


sub isNull
{
	my ( $work, $key, $errMsg) = @_;

	if( defined $work->{$key} ){
		if( $work->{$key} ne '' ){
			return 0; # nullでない
		}
	}
	
	$gTk{'window'}->messageBox(
		-type => "ok",
		-icon => "info",
		-title => "mergecsv",
		-message => $errMsg
		);
	return 1; # null
}

sub fileSelector
{
	my ( $file, $pattern, $mode, $is_only_file) = @_;
	my $suffix = [ # ファイル拡張子の設定
		[ $pattern,     [ '.'.$pattern ] ],
		[ 'All Files', [ '*'    ] ],
		];
	# ファイル選択ウインドウ
	# なぜかutf8のまま引渡しできる。。。？？？
	# my $path = encode('cp932', $gCWD);
	my $cwd = dirname( $$file ) . '/';
	$cwd =~ tr/\//\\/;
	#print "cwd[$cwd]\n";
	my $new_file = undef;
	if( $mode eq 'open' ){
		$new_file = $gTk{'window'}->getOpenFile( -filetypes  => $suffix,
												 -initialdir => $cwd,
												 -title      => '読み込むファイルを指定せんかい！'
												 );
	}
	else{
		$new_file = $gTk{'window'}->getSaveFile(-filetypes  => $suffix,
												-initialdir => $cwd,
												-title      => '保存するファイルを指定せんかい！'
												);
	}
	if( $new_file ){
		$new_file =~ s/\//\\/g;
		if( defined $is_only_file ){
			$new_file =~ s/^(.+\\)(.+?)$/$2/g;
		}
		$$file = $new_file;
	}
}

sub pathSelector
{
	my ($path) = @_;

	my $dir = encode('cp932', $$path );
	$dir =~ s/\//\\/g;
	my $folder = Win32::GUI::BrowseForFolder( -directory => $dir );
	return unless ($folder);
	$$path = decode('cp932', $folder);
	print ("path[$$path]\n");
}


sub mainProcess
{
	my $argv = shift;
	
	my $patternFile  = $argv->[0];
	my $targetDir = $argv->[1];

	setProgress(10);
	# 出力ファイル
	my $outFile = $patternFile;
	$outFile =~ s/\e|\*|\n//g;
	$outFile =~ s/\./_/g;
	$outFile = $outFile.'.xls';
	if( exists $gOptions->{'o'} ){
		$outFile = $gOptions->{'o'};
	}
	
	my $actualFileList = getFileList($targetDir, $patternFile);
	my $actualFilesValue = getFilesValue($targetDir, $actualFileList);
	setProgress(50);
	
	my $expectedFilesValue = undef;
	if( exists $gOptions->{'expected_path'} ){
		my $expectedFileList = getFileList($gOptions->{'expected_path'}, $patternFile);
		$expectedFilesValue = getFilesValue($gOptions->{'expected_path'}, $expectedFileList);
	}
	elsif( exists $gOptions->{'expected_xls_path'} ){
		# エクセルからexpectedを取得する
		$expectedFilesValue = getFilesValueFromXLS( ".", $gOptions->{'expected_xls_path'});
	}
	setProgress(70);
	
	my $ssdata = makeSpreadsheetData($actualFilesValue, $expectedFilesValue);
	exportExcel($outFile,$ssdata);
	sys_print("complate.\n");
	setProgress(100);
}

sub getOptions
{
	my ($argv) = @_;
	my %options = ();
	my @newAragv = ();
	for(my $i=0; $i< @{$argv}; $i++){
		my $key = decode('cp932', $argv->[$i]);
		if( $key =~ /^-(expected|xls|o)$/ ){
			$options{$1} = decode('cp932', $argv->[$i+1]);
			$i++;
		}
		elsif( $key =~ /^-(dbg|nogui)$/ ){
			$options{$1} = 1;
		}
		elsif( $key =~ /^-/ ){
			die "illigal parameter with options ($key)";
		}
		else{
			push @newAragv, $key;
		}
	}
	return (\@newAragv, \%options);
}

sub dbg_print
{
	my ($string) = @_;
	if( defined $gOptions->{'nogui'} ){
		if( defined $gOptions->{'dbg'} ){
			print encode('cp932', $string);
		}
	}
	else{
		$gTk{'log'}->insert('end', $string);
		$gTk{'log'}->see( 'end' );
	}
}

sub _print_
{
	my ($string, $col) = @_;
	
	if( defined $gOptions->{'nogui'} ){
		print encode('cp932', $string);
	}
	else{
		$gTk{'log'}->insert('end', $string);
		if( defined $col ){
			$gTk{'log'}->itemconfigure('end', -fg=>$col);
		}
		$gTk{'log'}->see( 'end' );
	}
}

sub err_print
{
	my ($string) = @_;
	_print_($string, '#ff0000');
}

sub sys_print
{
	my ($string) = @_;
	_print_($string);
}

# ファイルリスト取得
sub getFileList
{
	my ($path, $patternFile) = @_;
	my $path_sjis = encode('cp932', $path);

	my $regex = qr/$patternFile/;
	
	my @outFileList = ();
	opendir(DIR, $path_sjis);
	my @files_sjis = readdir(DIR);
	closedir(DIR);
	
	foreach my $file_sjis (@files_sjis){
		my $file = decode('cp932', $file_sjis);
		next if $file =~ /^\.{1,2}$/; # '.','..'は読み飛ばし
		my $filePath = $path.'\\'.$file;
		my $filePath_sjis = encode('cp932', $filePath);
		if( -d $filePath_sjis ){
			my $tempFileList = getFileList( $filePath, $patternFile);
			push @outFileList,@{$tempFileList};
			next;
		}
		if( $file =~ /$regex$/ ){
			push @outFileList, $filePath;
			dbg_print "getFileList() file[$filePath]\n";
		}
	}
	return \@outFileList;
}

sub getFileValue
{
	my ($file) = @_;
	my $file_sjis = encode('cp932',$file);
	open(IN ,"<$file_sjis") or die("error[$file_sjis]:$!");
	my @array = <IN> ;
	close IN;
	my %fileValue = ();
	my %keyHash = ();
	my $line_count=0;
	foreach my $line (@array){
		$line_count++;
		chomp $line;
		my $tmp = $line;
		$tmp =~ s/(?:\x0D\x0A|[\x0D\x0A])?$/,/;
		my @values = map {/^\"(.*)\"$/ ? scalar($_ = $1, s/\"\"/\"/g, $_) : $_}
		                  ($tmp =~ /(\"[^\"]*(?:\"\"[^\"]*)*\"|[^,]*),/g);
		push @{$fileValue{'keys'}} ,$values[0];
		# キー重複確認
		if( exists $keyHash{$values[0]} ){
			err_print "$file:$line_count,$keyHash{$values[0]}:Error [$values[0]]multiple keys.\n";
		}
		else{
			# 重複してない場合だけ値を保持(なんで重複キー発見以降のvaluesは無視する)
			$keyHash{$values[0]} = $line_count;
			for(my $i=1; $i<scalar(@values); $i++){
				push @{$fileValue{'data'}{$values[0]}} , $values[$i];
			}
		}
		
		# 最大要素数を取得
		my $count = scalar(@values) - 1; # key 要素除く
		if( defined $fileValue{'element_count'} ){
			if( $count > $fileValue{'element_count'} ){
				$fileValue{'element_count'} = $count;
			}
		}
		else{
			$fileValue{'element_count'} = $count;
		}
	}
	return (\%fileValue);
}

sub getFilesValue
{
	my ($basePath, $fileList) = @_;
	
	my %filesValue = ();
	$filesValue{'base_path'} = $basePath;
	foreach my $file (@{$fileList}){
		$filesValue{'files'}{$file} = getFileValue($file);
		dbg_print "[$file] elemnt count[$filesValue{'files'}{$file}{'element_count'}]\n";
	}
	return \%filesValue;
}

# $filesValue->{'base_path'}                             : ベースパス
# $filesValue->{'files'}{$ファイル名}{'keys'}[]          : 要素名リスト
#                                    {'data'}{$要素名}[] : 要素
#                                    {'element_count'}   : 最大要素数
sub getFilesValueFromXLS
{
	my ($basePath,$xlsfile) = @_;
	
	my %filesValue = ();
	$filesValue{'base_path'} = $basePath;
	%{$filesValue{'files'}} = ();

#	print "[$basePath][$xlsfile]\n";
	my $file_sjis = encode('cp932',$xlsfile);
	my $book = Spreadsheet::ParseExcel::Workbook->Parse($file_sjis);
	if( !$book ){
		die "[$xlsfile] not found.\n";
	}
#	print "book[$book]\n";
	my $sheet = $book->worksheet('expected');
	if( !$sheet ){
		die "[expected] sheet not found.\n";
	}
	
	# 要素名取得
	my @keys = ();
	for(my $row=1;$row<=$sheet->{MaxRow};$row++){
		my $cell = $sheet->get_cell($row,0);
		if( !defined $cell){
			last;
		}
		push @keys, $cell->value;
	}

	# ファイル名+要素取得
	# 未実装：フォーマット異常を検出する必要がある
	my @fileList = ();
	my $element_count = 0;
	my $fileName = undef;
	for(my $col=1;$col<=$sheet->{MaxCol};$col++){
		my $cell = $sheet->get_cell(0,$col);
		if( !defined $cell){
			last;
		}
		if( $cell->value ){
			# 新ファイル
			$fileName = $basePath."\\".$cell->value;
			push @fileList, $fileName;
			$element_count = 0;
			
			%{$filesValue{'files'}{$fileName}} = ();
			my $fileValue = $filesValue{'files'}{$fileName};
			push @{$fileValue->{'keys'}}, @keys;
			%{$fileValue->{'data'}} = ();
		}
		my $fileValue = $filesValue{'files'}{$fileName};
		my $data = $fileValue->{'data'};
		for(my $row=1;$row<=$sheet->{MaxRow};$row++){
			my $cell = $sheet->get_cell($row,$col);
			if( !defined $cell){
				last;
			}
			#print "$fileName\:$row\:$col=$keys[$row-1]\[$element_count\],".$cell->value."\n";
			$data->{$keys[$row-1]}[$element_count] = $cell->value;
		}
		$element_count++;
		$fileValue->{'element_count'} = $element_count;
	}
	return \%filesValue;
}

#fileValue{'keys'}[]        : 要素名リスト
#         {'data'}{$要素名} : 要素
#         {'element_count'} : 最大要素数
sub isSameFileValue
{
	my ($baseFileValue, $fileValue) = @_;

	# 要素数チェック
	if( $baseFileValue->{'element_count'} != $fileValue->{'element_count'} ){
		return 0;
	}
	# 要素名数チェック
	if( scalar(@{$baseFileValue->{'keys'}}) != scalar(@{$fileValue->{'keys'}}) ){
		return 0;
	}
	# 要素名チェック
	for(my $i=0; $i<scalar(@{$baseFileValue->{'keys'}}); $i++){
		if( $baseFileValue->{'keys'}[$i] ne $fileValue->{'keys'}[$i] ){
			return 0;
		}
	}
	
	return 1;
}


#            |  title_label
# -----------+--------------
# item_label |  values
#            |
sub makeSpreadsheetData
{
	my ($actualFilesValue, $expectedFilesValue) = @_;

	# actual file がない
	if( !exists $actualFilesValue->{'files'} ){
		die "[$actualFilesValue->{'base_path'}]actual file is not found.\n";
	}
	
	my %ssd = ();
	my $isUseExpected = 0;
	if( defined $expectedFilesValue ){
		$isUseExpected = 1;
		# expected file がない
		if( !exists $expectedFilesValue->{'files'} ){
			die "[$expectedFilesValue->{'base_path'}] expected file is not found.\n";
		}
	}
	
	# $filesValue->{'base_path'}                             : ベースパス
	# $filesValue->{'files'}{$ファイル名}{'keys'}[]          : 要素名リスト
	#                                    {'data'}{$要素名}[] : 要素
	#                                    {'element_count'}   : 最大要素数
	my $baseFileValue = undef;
	my $fileCnt = 0;
	foreach my $fileName ( keys %{$actualFilesValue->{'files'}} ){
		my $relativeFileName = $fileName;
		my $reg = quotemeta $actualFilesValue->{'base_path'};
		$relativeFileName =~ s/^$reg\\(.+?)$/$1/;
		
		my $fileValue = $actualFilesValue->{'files'}{$fileName};
		# 1件目をベースとするため、2件目以降はフォーマット確認をする
		if( defined $baseFileValue ){
			# 2件目以降
			# keys と 要素数があっているか確認
			if( !isSameFileValue($baseFileValue, $fileValue) ){
				err_print "[$fileName]Error:base file not same format. skip...\n";
				next;
			}
		}
		else{
			# 1件目
			# ベースファイルを先頭ファイルにする
			$baseFileValue = $fileValue;
			# item_label 作成
			foreach my $key (@{$fileValue->{'keys'}}){
				push @{$ssd{'item_label'}}, $key;
				#print "title_label=$key\n";
			}
			@{$ssd{'values'}} = ();
			@{$ssd{'values'}[0]} = ();
		}
		# title_label 作成
		push @{$ssd{'title_label'}}, $relativeFileName;
		for(my $i=0; $i<$fileValue->{'element_count'}-1;$i++){
			push @{$ssd{'title_label'}}, ''; # 要素数分進める
		}
		my $offset = $fileValue->{'element_count'} * 1 * $fileCnt;
		if( $isUseExpected ){
			# 想定結果がある場合(要素数x2の領域が足される)
			my $expectedFileName = "$expectedFilesValue->{'base_path'}\\$relativeFileName";
			my $expectedFileValue = $expectedFilesValue->{'files'}{$expectedFileName};
			# 結果ファイルが存在しているか確認
			if( !defined $expectedFileValue ){
				die "[$expectedFileName]expected file is not found.\n";
			}
			# フォーマットがactual と同じか確認
			if( !isSameFileValue($baseFileValue, $expectedFileValue) ){
				err_print "[$expectedFileName]Error:actual base file not same format. skip...\n";
				next;
			}
			
			#print "expectedFileName[$expectedFileName]\n";
			push @{$ssd{'title_label'}}, "expected\n$expectedFileName";
			for(my $i=0; $i<$expectedFileValue->{'element_count'}-1;$i++){
				push @{$ssd{'title_label'}}, ''; # 要素数分進める
			}
			push @{$ssd{'title_label'}}, "合否";
			for(my $i=0; $i<$expectedFileValue->{'element_count'}-1;$i++){
				push @{$ssd{'title_label'}}, ''; # 要素数分進める
			}
			$offset = $expectedFileValue->{'element_count'} * 3 * $fileCnt;
		}
		
		# values 作成
		my $row = 0;
		my $base_col= $offset;
		foreach my $key (@{$fileValue->{'keys'}}){
			my $cnt = 0;
			foreach my $value (@{$fileValue->{'data'}{$key}}){
				$ssd{'values'}[$row][$base_col+$cnt] = $value;
				$cnt++;
			}
			for(;$cnt<$fileValue->{'element_count'};$cnt++){ # 要素が少ない場合、'―'文字で埋めておく
				$ssd{'values'}[$row][$base_col+$cnt] = '―';
			}
			$row++;
		}
		
		if( $isUseExpected ){
			# 想定結果がある場合
			my $expectedFileName = "$expectedFilesValue->{'base_path'}\\$relativeFileName";
			my $expectedFileValue = $expectedFilesValue->{'files'}{$expectedFileName};
			my $row = 0;
			my $base_col= $offset + $fileValue->{'element_count'};
			foreach my $key (@{$expectedFileValue->{'keys'}}){
				my $cnt = 0;
				# 結果要素
				foreach my $value (@{$expectedFileValue->{'data'}{$key}}){
					$ssd{'values'}[$row][$base_col+$cnt] = $value;
					if ($fileValue->{'data'}{$key}[$cnt] ne $value) {
						printf("$expectedFileName:[$cnt][$key]:actual[$fileValue->{'data'}{$key}[$cnt]]:expected[$value]\n");
					}
					$cnt++;
				}
				for(;$cnt<$fileValue->{'element_count'};$cnt++){ # 要素が少ない場合、'―'文字で埋めておく
					$ssd{'values'}[$row][$base_col+$cnt] = '―';
				}
				
				# 合否判定
				for(my $i=0; $i<$fileValue->{'element_count'}; $i++){
					my $drow = $row + 2;
					my $act = excel_num2col(1+$offset+$i);
					my $exp = excel_num2col(1+$offset+$fileValue->{'element_count'}+$i);
					my $cmd = "=IF(EXACT($act$drow,$exp$drow)=TRUE,\"○\",\"×\")";
					$ssd{'values'}[$row][$base_col+$cnt] = $cmd;
					$cnt++;
				}
				$row++;
			}
		}
		$fileCnt++;
	}
	return \%ssd;
}


#            |  title_label
# -----------+--------------
# item_label |  values
#            |
#
# $ssd->{'title_label'}[] : タイトルリスト
# $ssd->{'item_label'}[] : 要素名リスト
# $ssd->{'values'}[][] : 要素
sub exportExcel
{
	my ($file,$ssd) = @_;

	dbg_print "exportExcel[$file]\n";
	my $file_sjis = encode('cp932', $file);
	my $workbook = Spreadsheet::WriteExcel->new($file_sjis); # Create a new Excel workbook
	my $worksheet = $workbook->add_worksheet(); # Add a worksheet
	{
		# title_label
		my $format = $workbook->add_format(); # Add and define a format
		$format->set_bg_color('silver');
		$format->set_bold();
		$format->set_align('left');
		$format->set_align('top');
		$format->set_text_wrap();
		$worksheet->set_column( 1, scalar(@{$ssd->{'title_label'}}), 20);
		for(my $x=0;$x<scalar(@{$ssd->{'title_label'}});$x++){
			$worksheet->write( 0, 1+$x, $ssd->{'title_label'}[$x], $format);
		}
	}
	{
		# item_label
		my $format = $workbook->add_format(); # Add and define a format
		#$format->set_bg_color('silver');
		$format->set_bold();
		$format->set_align('left');
		$format->set_align('top');
		$format->set_text_wrap();
		$worksheet->set_column( 0, 0, 10);
		for(my $y=0;$y<scalar(@{$ssd->{'item_label'}});$y++){
			$worksheet->write( 1+$y, 0, $ssd->{'item_label'}[$y], $format);
		}
	}
	{
		# values
		my $format = $workbook->add_format(); # Add and define a format
		$format->set_align('left');
		$format->set_align('top');

# 		if(0){
# 			my $y = 1;
# 			foreach my $values (@{$ssd->{'values'}}){
# 				my $x = 1;
# 				foreach my $value (@{$values}){
# 					$worksheet->write( $y, $x, $value, $format);
# 					$x++;
# 				}
# 				$y++;
# 			}
# 		}else{
			for(my $y=0;$y<scalar(@{$ssd->{'values'}});$y++){
				for(my $x=0;$x<scalar(@{$ssd->{'values'}[$y]});$x++){
					$worksheet->write( 1+$y, 1+$x, $ssd->{'values'}[$y][$x], $format);
				}
			}
#		}
	}
	$workbook->close;
}

# num to A-Z for excel.
sub excel_num2col
{
	my ($bb) = @_;
	
	my @dst = ();
	do {
		my $mod = ($bb % 26);
		push @dst, unpack("C", 'A') + $mod;
		$bb = int( $bb / 26) - 1;
	}while( $bb >= 0 );

#	return join '',@dst;
# #	printf("[%s][%s]\n",chr($dst[0]),chr($dst[1]));
 	my $line = '';
 	foreach my $str (@dst){
 		$line = chr($str).$line;
 	}
# #	printf("sub : $line\n");
	return $line;
}

#----------------------------------------------------------------
# A-Z to num for excel.
sub excel_col2num
{
	my ($a) = @_;
	my @strs = split //, $a;
	my $digit = 0;
	my $ans = 0;
	foreach my $num (@strs){
		my $tmp = unpack("C", $num) - unpack("C", 'A') + 1;
		$ans = ($ans * 26);
#		print "$digit : \($num == $tmp\) += $ans\n";
		$ans += $tmp;
		$digit++;
	}
	return $ans - 1;
#	print "ans = $ans\n";
}

#----------------------------------------------------------------
# col = col + col
sub excel_add_col
{
	my ($a,$b) = @_;

	my $aa = excel_col2num($a);
	my $bb = excel_col2num($b);
	my $ans = $aa + $bb;
	$ans = excel_num2col($ans);
	return $ans;
}

#----------------------------------------------------------------
# col = col + num
sub excel_add_num
{
	my ($a,$b) = @_;

	my $aa = excel_col2num($a);
	my $ans = $aa += $b;
	$ans = excel_num2col($ans);
	return $ans;
}

