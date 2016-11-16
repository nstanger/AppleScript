#!/bin/tcsh

set dir=$1

foreach infile ($dir/*.applescript)
    set outfile=`basename $infile | sed -e 's/applescript$/scpt/'`
    osacompile -o "$outfile" "$infile"
    mv "$outfile" "$HOME/Library/Scripts/Applications/$dir"
end
