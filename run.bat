@echo off
BREAK=ON

set oldclasspath=%CLASSPATH%

set CLASSPATH=..\ConsoleOne.jar
set CLASSPATH=%CLASSPATH%;..\ConsoleOneCore.jar
set CLASSPATH=%CLASSPATH%;..\ConsoleOneExt\collections.jar
set CLASSPATH=%CLASSPATH%;..\ConsoleOneExt\jgl3.1.0.jar
set CLASSPATH=%CLASSPATH%;..\ConsoleOneExt\jh.jar
set CLASSPATH=%CLASSPATH%;..\ConsoleOneExt\Widgets.jar
set CLASSPATH=%CLASSPATH%;..\help
set CLASSPATH=%CLASSPATH%;..\ConsoleOneExt\njha.jar
set CLASSPATH=%CLASSPATH%;..\ConsoleOneExt\jclient.jar
set CLASSPATH=%CLASSPATH%;..\ConsoleOneExt\njclv2.jar
set CLASSPATH=%CLASSPATH%;..\ConsoleOneExt\NWFS.jar
set CLASSPATH=%CLASSPATH%;..\ConsoleOneExt\JReportBeans.zip

set BOOT_CLASSPATH=..\jre\lib\rt.jar;..\jre\lib\i18n.jar;..\jre\lib\charsets.jar


@echo on
..\jre\bin\java -Xbootclasspath:%BOOT_CLASSPATH% -Djava.ext.dirs= -Djava.security.manager -Djava.security.policy=ConsoleOne.policy -noverify -classpath %CLASSPATH% com.novell.application.console.shell.Console %1 %2 %3 %4
@echo off
set CLASSPATH=%oldclasspath%
set oldclasspath=
