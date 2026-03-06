@echo off
:: [v11.5] ApiExcelExporter 통합 실행 배치 (Portable Tools & Memory 최적화)
chcp 65001
cls

:: [경로 설정] 현재 배치파일 위치의 tools 폴더를 기준으로 함
set TOOL_DIR=%~dp0tools
set JAVA_BIN="%TOOL_DIR%\jdk\bin\java.exe"
set GIT_BIN="%TOOL_DIR%\git\bin\git.exe"

:: [환경변수] 시스템 PATH보다 tools 내 경로를 우선하도록 설정
set PATH=%TOOL_DIR%\jdk\bin;%TOOL_DIR%\git\bin;%PATH%

:: [라이브러리 및 클래스 경로]
set PROJECT_ROOT=.
set CLASSPATH=%PROJECT_ROOT%\target\classes;%PROJECT_ROOT%\lib\*

:: [JVM 메모리 옵션] i9-13900 32GB 환경 최적화
set JAVA_OPTS=-Xms4g -Xmx8g -XX:+UseG1GC -XX:+HeapDumpOnOutOfMemoryError

echo ===============================================================
echo  [RUN] ApiExcelExporter v11.5 실행 (Office: i9-13900)
echo  [TOOL] Java Path: %JAVA_BIN%
echo  [TOOL] Git Path: %GIT_BIN%
echo  [MEM] JVM Heap: Initial 4GB / Max 8GB
echo ===============================================================

%JAVA_BIN% %JAVA_OPTS% -Dfile.encoding=UTF-8 -cp "%CLASSPATH%" com.baek.ApiExcelExporter

echo.
echo 작업이 완료되었습니다. 로그를 확인하세요.
pause