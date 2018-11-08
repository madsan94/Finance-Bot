#!/bin/sh
while IFS='' read -r line || [[ -n "$line" ]]; do
    python3 TestBot.py $line
    echo $line
done < "$1"
