#!/bin/bash

if [ "$1" = "--clean" ]
then
    git clean -fXd
    exit 0
fi

find . -path ./.git -prune -o -name "*.applescript" -execdir osacompile -o "{}.scpt" "{}" \;
find . -path ./.git -prune -o -name "*.scpt" -execdir rename --force 's/\.applescript\.scpt/.scpt/' "{}" \;
rsync -av --exclude="*.applescript" Applications $HOME/Library/Scripts
rsync -av --exclude="*.applescript" "Script Libraries" $HOME/Library
