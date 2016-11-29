#!/bin/tcsh

set dir=$1

foreach infile ($dir/*.applescript)
    set outfile=`basename $infile | sed -e 's/applescript$/scpt/'`
    osacompile -o "$outfile" "$infile"
    mkdir -p "$HOME/Library/Scripts/Applications/$dir"
    mv "$outfile" "$HOME/Library/Scripts/Applications/$dir"
end
