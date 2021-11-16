#!/bin/bash

if [ "$1" = "--clean" ]
then
    /opt/local/bin/git clean -fXd
    exit 0
fi
