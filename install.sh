#!/bin/bash

if [ "$1" = "--clean" ]
then
    /opt/local/bin/git clean -fXd
    exit 0
fi

/usr/local/opt/findutils/libexec/gnubin/find . -path ./.git -prune -o -name "*.applescript" -execdir osacompile -o "{}.scpt" "{}" \;
/usr/local/opt/findutils/libexec/gnubin/find . -path ./.git -prune -o -name "*.scpt" -execdir rename --force 's/\.applescript\.scpt/.scpt/' "{}" \;
/usr/local/bin/rsync -av --exclude="*.applescript" Applications $HOME/Library/Scripts
/usr/local/bin/rsync -av --exclude="*.applescript" "Script Libraries" $HOME/Library
