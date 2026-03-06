@echo off
:: [v5.5] WhatapApiCounter 단독 실행 배치 (Portable Tools & Memory 최적화)
chcp 65001
cls

:: [경로 설정]
set TOOL_DIR=%~dp0tools
set JAVA_BIN="%TOOL_DIR%\jdk\bin\java.exe"

:: [환경변수]
set PATH=%TOOL_DIR%\jdk\bin;%PATH%

:: [라이브러리 경로]
set PROJECT_ROOT=.
set CLASSPATH=%PROJECT_ROOT%\target\classes;%PROJECT_ROOT%\lib\*

:: [JVM 메모리 옵션] 단독 실행용 (4GB 상한)
set JAVA_OPTS=-Xms2g -Xmx4g -XX:+UseG1GC

echo ===============================================================
echo  [RUN] WhatapApiCounter v5.5 단독 실행
echo  [TOOL] Java Path: %JAVA_BIN%
echo  [MEM] JVM Heap: Initial 2GB / Max 4GB
echo ===============================================================

%JAVA_BIN% %JAVA_OPTS% -Dfile.encoding=UTF-8 -cp "%CLASSPATH%" com.baek.WhatapApiCounter

echo.
echo 와탭 통계 추출이 완료되었습니다.
pause