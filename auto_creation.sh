#!/bin/bash

env_info="MINGW32_NT"
exec_env=`uname`

if [ ${exec_env:0:10} != $env_info ]; then
	echo "Please install Git on Windows."
	exit 1
fi

for dir in A0*
do
	if [ -d $dir ]; then
		ls $dir/*.bas >/dev/null 2>&1
		if [ $? -ne 0 ]; then
			echo "$dir: Export bas file(s)"
			cscript.exe ./vba_exporter.vbs $dir
		else
			echo "$dir: Skip exporting"
		fi
fi
done
exit 0
