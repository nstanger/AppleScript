#!/bin/bash

if [ "$1" = "--clean" ]
then
    /opt/local/bin/git clean -fXd
    exit 0
fi

/opt/local/bin/gfind . -path ./.git -prune -o -name "*.applescript" -execdir osacompile -o "{}.scpt" "{}" \;
/opt/local/bin/gfind . -path ./.git -prune -o -name "*.scpt" -execdir rename -force 's/\.applescript\.scpt/.scpt/' "{}" \;
/opt/local/bin/rsync -av --exclude="*.applescript" Applications $HOME/Library/Scripts
/opt/local/bin/rsync -av --exclude="*.applescript" "Script Libraries" $HOME/Library
