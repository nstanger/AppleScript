#!/bin/bash

/opt/local/bin/gfind . -name "*.applescript" -execdir osacompile -o "{}.scpt" "{}" \;
/opt/local/bin/gfind . -name "*.scpt" -execdir rename -force 's/\.applescript\.scpt/.scpt/' "{}" \;
/opt/local/bin/rsync -av --exclude="*.applescript" Applications $HOME/Library/Scripts
/opt/local/bin/rsync -av --exclude="*.applescript" "Script Libraries" $HOME/Library

if [ "$1" = "--clean" ]
then
    /opt/local/bin/gfind -name "*.scpt" -delete
fi
