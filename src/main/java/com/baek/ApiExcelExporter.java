import com.github.javaparser.StaticJavaParser;
import com.github.javaparser.ast.CompilationUnit;
import com.github.javaparser.ast.body.ClassOrInterfaceDeclaration;
import com.github.javaparser.ast.body.MethodDeclaration;
import com.github.javaparser.ast.expr.*;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.regex.*;
import java.util.stream.Collectors;

/**
 * 프로젝트명: ApiExcelExporter (Bitbucket 관리형)
 * Version: 11.2 (컬럼 확장 및 Ivory 테마 적용)
 * 반영사항:
 * 1. [지적 반영] 주요 변수 기본값 비우기("")를 통한 보안 강화 (정보 노출 방지)
 * 2. [기능 추가] 엑셀 우측에 관리용 컬럼 5개 추가 (조치예정일자, 조치일자, 관련티켓, 조치담당자, 비고)
 * 3. [디자인] 신규 컬럼 5개에 '아이보리(Lemon Chiffon)' 헤더 색상 적용
 * 4. [로직 보존] i9-13900 최적화 parallelStream 및 Whatap 통합 로직 유지 (v11.1 베이스)
 */
public class ApiExcelExporter {

    // ==========================================================================================
    // [ 1. 내부 기본 설정부 ] - config.properties를 반드시 작성하세요.
    // ==========================================================================================

    /** [핵심변수 1] 레파지토리 이름 */
    private static String REPO_NAME = "";

    /** [핵심변수 2] 기본 도메인 주소  */
    private static String DOMAIN = "";

    /** [핵심변수 3] 분석할 Java 소스 로컬 절대 경로 */
    private static String ROOT_PATH = "";

    /** [핵심변수 4] 결과 저장 디렉토리 물리적 경로 */
    private static String OUTPUT_DIR = "";

    /** [핵심변수 5] Git 실행 경로 (환경변수 미등록 대비) */
    private static String GIT_BIN_PATH = "git";

    /** [상태변수] 외부 설정 파일(config.properties) 로드 성공 여부 */
    private static boolean isConfigLoaded = false;

    // ==========================================================================================

    private static final List<String> MAPPING_ANNS = Arrays.asList("RequestMapping", "GetMapping", "PostMapping", "PutMapping", "DeleteMapping", "PatchMapping");
    private static final List<String> RUNTIME_LOGS = Collections.synchronizedList(new ArrayList<>());
    private static String logPath = "";
    private static final AtomicInteger PROCESSED_COUNT = new AtomicInteger(0);

    public static void main(String[] args) {
        // 1. 설정 로드
        loadExternalConfig();

        if (OUTPUT_DIR.isEmpty()) {
            System.err.println("[ERROR] OUTPUT_DIR이 설정되지 않았습니다. config.properties를 확인하세요.");
            return;
        }

        File dir = new File(OUTPUT_DIR);
        if (!dir.exists()) dir.mkdirs();

        long startTime = System.currentTimeMillis();
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));

        System.out.println("===============================================================");
        System.out.println("[START] " + REPO_NAME + " API 추출 및 Whatap 통합 시작 (v11.2)");
        if (isConfigLoaded) System.out.println("[INFO] 설정 파일 로드 성공: 외부 config.properties 사용.");
        else System.out.println("[INFO] 설정 파일 로드 실패: 기본값이 비어 있어 분석이 진행되지 않을 수 있습니다.");
        System.out.println("[INFO] 현재 분석 경로: " + ROOT_PATH);
        System.out.println("[INFO] Git 실행 경로: " + GIT_BIN_PATH);
        System.out.println("===============================================================");

        // 2. Whatap 통계 모듈 가동
        System.out.println("[INFO] 와탭 통계 수집 및 엑셀 리포트 생성 중...");
        Map<String, long[]> whatapStats = WhatapApiCounter.getApiStats();
        WhatapApiCounter.generateExcelReport(timestamp);
        System.out.println("[INFO] 와탭 데이터 확보 완료 (총 " + whatapStats.size() + "개 트랜잭션)");

        List<ApiInfo> allApiList = Collections.synchronizedList(new ArrayList<>());
        int totalFiles = 0;

        try {
            Path rootPathObj = Paths.get(ROOT_PATH);
            List<Path> controllerFiles = Files.walk(rootPathObj)
                    .filter(p -> p.toString().endsWith(".java") &&
                            (p.toString().contains("Controller") || p.toString().contains("Conrtoller")))
                    .collect(Collectors.toList());

            totalFiles = controllerFiles.size();
            final int total = totalFiles;

            controllerFiles.parallelStream().forEach(file -> {
                String relativePath = rootPathObj.relativize(file).toString();
                int current = PROCESSED_COUNT.incrementAndGet();

                List<String[]> gitHistories = getRecentGitHistories(relativePath, ROOT_PATH, 3);

                StringBuilder fileLog = new StringBuilder();
                fileLog.append(String.format("\n[%d/%d] 분석: %s", current, total, file.getFileName()));
                fileLog.append(String.format(" (최신커밋: %s | %s)", gitHistories.get(0)[0], gitHistories.get(0)[1]));

                allApiList.addAll(extractApisHybrid(file, relativePath, gitHistories, fileLog));
                System.out.print(fileLog.toString());
                synchronized (RUNTIME_LOGS) { RUNTIME_LOGS.add(fileLog.toString()); }
            });
        } catch (Exception e) {
            addExceptionLog("디렉토리 탐색 오류", e);
            return;
        }

        allApiList.sort(Comparator.comparing(ApiInfo::getApiPath));

        // [지적 반영] 파일명 규격화 (v11.2 업데이트) [cite: 2026-02-23]
        String baseFileName = String.format("API 현황 추출결과_(v11.2)_(%s)_(컨트롤러  %d개 & API %d개)_(%s)",
                REPO_NAME, totalFiles, allApiList.size(), timestamp);

        logPath = OUTPUT_DIR + File.separator + baseFileName + ".log";
        saveInitialLogsToPath();

        File finalExcelFile = new File(OUTPUT_DIR, baseFileName + ".xlsx");

        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(finalExcelFile)) {

            String sheetName = "API분석_" + REPO_NAME;
            if (sheetName.length() > 31) sheetName = sheetName.substring(0, 31);
            Sheet sheet = workbook.createSheet(sheetName);
            CreationHelper helper = workbook.getCreationHelper();

            // 기존 테마 색상 정의
            CellStyle greyH = createStyle(workbook, IndexedColors.GREY_25_PERCENT.getIndex(), true, true);
            CellStyle yellowH = createStyle(workbook, IndexedColors.YELLOW.getIndex(), true, true);
            CellStyle orangeH = createStyle(workbook, IndexedColors.ORANGE.getIndex(), true, true);
            CellStyle greenH = createStyle(workbook, IndexedColors.LIGHT_GREEN.getIndex(), true, true);
            CellStyle blueH = createStyle(workbook, IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex(), true, true);

            // [신규] 아이보리(Lemon Chiffon) 테마 정의 [cite: 2026-02-23]
            CellStyle ivoryH = createStyle(workbook, IndexedColors.LEMON_CHIFFON.getIndex(), true, true);

            CellStyle centerD = createStyle(workbook, null, false, true);
            CellStyle leftD = createStyle(workbook, null, false, false);
            CellStyle numD = workbook.createCellStyle(); numD.setDataFormat(workbook.createDataFormat().getFormat("#,##0"));
            numD.setBorderBottom(BorderStyle.THIN); numD.setBorderTop(BorderStyle.THIN); numD.setBorderLeft(BorderStyle.THIN); numD.setBorderRight(BorderStyle.THIN);

            CellStyle linkD = createStyle(workbook, null, false, false);
            Font linkFont = workbook.createFont(); linkFont.setColor(IndexedColors.BLUE.getIndex()); linkFont.setUnderline(Font.U_SINGLE);
            linkD.setFont(linkFont);
            CellStyle depColumnStyle = createStyle(workbook, IndexedColors.GREY_25_PERCENT.getIndex(), false, true);

            sheet.createFreezePane(3, 1);

            // [기능 추가] 헤더 확장 (총 27개 컬럼) [cite: 2026-02-23]
            String[] headers = {"순번","레파지토리","API 경로","전체 URL","repository path","컨트롤러명","호출메소드",
                    "Deprecated","커밋일자1","커밋터1","코멘트1","커밋일자2","커밋터2","코멘트2","커밋일자3","커밋터3","코멘트3",
                    "호출건수(APM추출필요)", "프로그램ID(필요시)", "담당자(필요시)", "미사용 의심건", "미사용 검토결과",
                    "조치예정일자", "조치일자", "관련티켓", "조치담당자", "비고"};

            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                if (i <= 3) cell.setCellStyle(greyH);
                else if (i <= 6) cell.setCellStyle(yellowH);
                else if (i <= 17) cell.setCellStyle(orangeH);
                else if (i <= 19) cell.setCellStyle(greenH);
                else if (i <= 21) cell.setCellStyle(blueH);
                else cell.setCellStyle(ivoryH); // 22~26번 인덱스: 아이보리 적용 [cite: 2026-02-23]
            }
            sheet.setAutoFilter(new CellRangeAddress(0, 0, 0, headers.length - 1));

            for (int i = 0; i < allApiList.size(); i++) {
                ApiInfo info = allApiList.get(i);
                Row row = sheet.createRow(i + 1);
                boolean isDep = "Y".equals(info.isDeprecated);
                String fullUrl = DOMAIN + info.apiPath;

                long totalCalls = 0;
                long[] rowStats = whatapStats.get(info.apiPath);
                if (rowStats != null) for (long count : rowStats) totalCalls += count;

                // 데이터 배열 확장 대응 (27개 항목)
                String[] data = {String.valueOf(i + 1), REPO_NAME, info.apiPath, fullUrl, info.repoPath,
                        info.controllerName, info.methodName, info.isDeprecated,
                        info.git1[0], info.git1[1], info.git1[2], info.git2[0], info.git2[1], info.git2[2],
                        info.git3[0], info.git3[1], info.git3[2],
                        String.valueOf(totalCalls), "", "", (totalCalls == 0 ? "O" : ""), "",
                        "", "", "", "", ""}; // 조치관련 5종 공백 초기화

                for (int j = 0; j < data.length; j++) {
                    Cell cell = row.createCell(j);
                    if (j == 17) {
                        cell.setCellValue(totalCalls);
                        cell.setCellStyle(numD);
                    } else {
                        cell.setCellValue(data[j]);
                        boolean isCenter = (j==0 || j==1 || j==5 || j==6 || j==7 || j==8 || j==9 || j==11 || j==12 || j==14 || j==15 || (j>=18));
                        if (j == 7 && isDep) cell.setCellStyle(depColumnStyle);
                        else if (j == 3) {
                            cell.setCellStyle(linkD);
                            try {
                                String encodedUrl = fullUrl.replace("{", "%7B").replace("}", "%7D");
                                Hyperlink link = helper.createHyperlink(HyperlinkType.URL);
                                link.setAddress(encodedUrl);
                                cell.setHyperlink(link);
                            } catch (Exception e) { }
                        } else cell.setCellStyle(isCenter ? centerD : leftD);
                    }
                }
            }

            sheet.setColumnWidth(2, 14500); sheet.setColumnWidth(3, 8500);
            sheet.setColumnWidth(4, 11500); sheet.setColumnWidth(5, 5500); sheet.setColumnWidth(6, 5500);
            for (int i = 8; i < headers.length; i++) sheet.setColumnWidth(i, 4200);

            workbook.write(fos);
            addLog("\n[SUCCESS] 통합 엑셀 저장 완료: " + finalExcelFile.getName());
        } catch (Exception e) { addExceptionLog("엑셀 저장 중 오류", e); }

        addLog("\n[FINISH] 작업 종료: " + (System.currentTimeMillis() - startTime) / 1000 + "초 소요");
    }

    private static void loadExternalConfig() {
        Properties prop = new Properties();
        File configFile = new File("config.properties");
        if (configFile.exists()) {
            try (InputStream is = new FileInputStream(configFile)) {
                prop.load(is);
                if (prop.getProperty("REPO_NAME") != null) REPO_NAME = prop.getProperty("REPO_NAME").trim();
                if (prop.getProperty("DOMAIN") != null) DOMAIN = prop.getProperty("DOMAIN").trim();
                if (prop.getProperty("ROOT_PATH") != null) ROOT_PATH = prop.getProperty("ROOT_PATH").trim();
                if (prop.getProperty("OUTPUT_DIR") != null) OUTPUT_DIR = prop.getProperty("OUTPUT_DIR").trim();
                if (prop.getProperty("GIT_BIN_PATH") != null) GIT_BIN_PATH = prop.getProperty("GIT_BIN_PATH").trim();
                isConfigLoaded = true;
            } catch (IOException ignored) {}
        }
    }

    private static List<ApiInfo> extractApisHybrid(Path path, String rel, List<String[]> git, StringBuilder log) {
        try { return extractWithJavaParser(path, rel, git, log); }
        catch (Exception e) {
            log.append("\n  ! [파싱 에러] ").append(path.getFileName()).append(" 사유: ").append(e.getMessage());
            return extractWithRegex(path, rel, git, log);
        }
    }

    private static List<ApiInfo> extractWithJavaParser(Path filePath, String relPath, List<String[]> git, StringBuilder log) throws Exception {
        List<ApiInfo> apis = new ArrayList<>();
        String rawCode = new String(Files.readAllBytes(filePath), StandardCharsets.UTF_8);
        CompilationUnit cu = StaticJavaParser.parse(rawCode);
        String classPath = "";
        Optional<ClassOrInterfaceDeclaration> mainClass = cu.findFirst(ClassOrInterfaceDeclaration.class);
        if (mainClass.isPresent()) {
            Optional<AnnotationExpr> classAnn = mainClass.get().getAnnotationByName("RequestMapping");
            if (classAnn.isPresent()) {
                List<String> cpList = getPathsFromAnn(classAnn.get());
                if (!cpList.isEmpty()) classPath = cpList.get(0).trim();
            }
        }
        for (MethodDeclaration method : cu.findAll(MethodDeclaration.class)) {
            for (String annName : MAPPING_ANNS) {
                Optional<AnnotationExpr> methodAnn = method.getAnnotationByName(annName);
                if (methodAnn.isPresent()) {
                    List<String> subPaths = getPathsFromAnn(methodAnn.get());
                    if (subPaths.isEmpty()) subPaths.add("");
                    for (String s : subPaths) {
                        String cp = classPath.startsWith("/") ? classPath : (classPath.isEmpty() ? "" : "/" + classPath);
                        String mp = s.trim().startsWith("/") ? s.trim() : (s.trim().isEmpty() ? "" : "/" + s.trim());
                        String finalPath = (cp + mp).replaceAll("/+", "/");
                        ApiInfo info = new ApiInfo();
                        info.apiPath = (finalPath.isEmpty() ? "/" : finalPath);
                        info.methodName = method.getNameAsString();
                        info.isDeprecated = method.isAnnotationPresent("Deprecated") ? "Y" : "N";
                        info.controllerName = filePath.getFileName().toString();
                        info.repoPath = (REPO_NAME + "/" + relPath).replace("\\", "/");
                        info.git1 = git.get(0); info.git2 = git.get(1); info.git3 = git.get(2);
                        apis.add(info);
                        log.append("\n    └ [Found] ").append(info.apiPath);
                    }
                }
            }
        }
        return apis;
    }

    private static List<ApiInfo> extractWithRegex(Path filePath, String relPath, List<String[]> git, StringBuilder log) {
        List<ApiInfo> apis = new ArrayList<>();
        try {
            String raw = new String(Files.readAllBytes(filePath), StandardCharsets.UTF_8);
            String clean = raw.replaceAll("(?s)/\\*.*?\\*/", " ").replaceAll("//.*", " ");
            String classPath = "";
            Matcher cm = Pattern.compile("@RequestMapping\\s*\\(\\s*(?:(?:value|path)\\s*=\\s*)?\"([^\"]+)\"").matcher(clean);
            if (cm.find()) classPath = cm.group(1).trim();
            Matcher mMatcher = Pattern.compile("@(GetMapping|PostMapping|RequestMapping|PutMapping|DeleteMapping|PatchMapping)\\s*\\((.*?)\\)", Pattern.DOTALL).matcher(raw);
            while (mMatcher.find()) {
                String params = mMatcher.group(2);
                Matcher mName = Pattern.compile("(?:public|private|protected)\\s+[\\w<>,\\s]+\\s+(\\w+)\\s*\\(").matcher(clean.substring(mMatcher.end(), Math.min(mMatcher.end() + 1000, clean.length())));
                if (mName.find()) {
                    Matcher p = Pattern.compile("\"([^\"]+)\"").matcher(params);
                    while (p.find()) {
                        String s = p.group(1).trim();
                        if (!s.contains("RequestMethod")) {
                            String cp = classPath.startsWith("/") ? classPath : (classPath.isEmpty() ? "" : "/" + classPath);
                            String mp = s.startsWith("/") ? s : (s.isEmpty() ? "" : "/" + s);
                            String finalPath = (cp + mp).replaceAll("/+", "/");
                            ApiInfo info = new ApiInfo();
                            info.apiPath = (finalPath.isEmpty() ? "/" : finalPath);
                            info.methodName = mName.group(1);
                            info.isDeprecated = clean.substring(Math.max(0, mMatcher.start() - 300), mMatcher.start()).contains("@Deprecated") ? "Y" : "N";
                            info.controllerName = filePath.getFileName().toString();
                            info.repoPath = (REPO_NAME + "/" + relPath).replace("\\", "/");
                            info.git1 = git.get(0); info.git2 = git.get(1); info.git3 = git.get(2);
                            apis.add(info);
                            log.append("\n    └ [Found-Regex] ").append(info.apiPath);
                        }
                    }
                }
            }
        } catch (Exception ignored) {}
        return apis;
    }

    private static List<String> getPathsFromAnn(AnnotationExpr ann) {
        List<String> paths = new ArrayList<>();
        Expression value = (ann instanceof SingleMemberAnnotationExpr) ? ((SingleMemberAnnotationExpr) ann).getMemberValue() : null;
        if (ann instanceof NormalAnnotationExpr) {
            value = ((NormalAnnotationExpr) ann).getPairs().stream()
                    .filter(p -> p.getNameAsString().equals("value") || p.getNameAsString().equals("path"))
                    .map(MemberValuePair::getValue).findFirst().orElse(null);
        }
        if (value instanceof ArrayInitializerExpr) {
            for (Expression expr : ((ArrayInitializerExpr) value).getValues()) if (expr instanceof StringLiteralExpr) paths.add(((StringLiteralExpr) expr).getValue());
        } else if (value instanceof StringLiteralExpr) paths.add(((StringLiteralExpr) value).getValue());
        return paths;
    }

    private static void addLog(String msg) {
        System.out.println(msg);
        if (logPath.isEmpty()) return;
        try (FileWriter fw = new FileWriter(logPath, true); PrintWriter pw = new PrintWriter(fw)) {
            pw.println(msg);
        } catch (IOException ignored) {}
    }

    private static void saveInitialLogsToPath() {
        try (FileWriter fw = new FileWriter(logPath, false); PrintWriter pw = new PrintWriter(fw)) {
            pw.println("===============================================================");
            pw.println("[START] " + REPO_NAME + " API 추출 및 Whatap 통합 시작 (v11.2)");
            pw.println("[INFO] 분석 경로: " + ROOT_PATH);
            pw.println("===============================================================");
            synchronized (RUNTIME_LOGS) { for (String l : RUNTIME_LOGS) pw.println(l); }
        } catch (IOException ignored) {}
    }

    private static void addExceptionLog(String title, Exception e) {
        StringWriter sw = new StringWriter(); e.printStackTrace(new PrintWriter(sw));
        addLog("\n[ERROR] " + title + "\n" + sw.toString());
    }

    private static List<String[]> getRecentGitHistories(String rel, String root, int c) {
        List<String[]> h = new ArrayList<>(); for (int i = 0; i < c; i++) h.add(new String[]{"-", "-", "No History"});
        try {
            Process p = new ProcessBuilder(GIT_BIN_PATH, "log", "-" + c, "--pretty=format:%as|%an|%s", "--", rel).directory(new File(root)).start();
            try (BufferedReader r = new BufferedReader(new InputStreamReader(p.getInputStream(), StandardCharsets.UTF_8))) {
                List<String> lines = new ArrayList<>(); String l; while ((l = r.readLine()) != null) lines.add(l);
                Collections.reverse(lines);
                for (int i = 0; i < lines.size(); i++) {
                    String[] parts = lines.get(i).split("\\|", 3);
                    h.set(i, new String[]{parts[0], parts[1], parts.length > 2 ? parts[2] : ""});
                }
            }
            p.waitFor();
        } catch (Exception ignored) {}
        return h;
    }

    private static CellStyle createStyle(Workbook wb, Short bg, boolean bold, boolean center) {
        CellStyle s = wb.createCellStyle();
        if (bg != null) { s.setFillForegroundColor(bg); s.setFillPattern(FillPatternType.SOLID_FOREGROUND); }
        s.setAlignment(center ? HorizontalAlignment.CENTER : HorizontalAlignment.LEFT);
        s.setVerticalAlignment(VerticalAlignment.CENTER);
        s.setBorderBottom(BorderStyle.THIN); s.setBorderTop(BorderStyle.THIN);
        s.setBorderLeft(BorderStyle.THIN); s.setBorderRight(BorderStyle.THIN);
        Font f = wb.createFont(); f.setBold(bold); s.setFont(f);
        return s;
    }

    static class ApiInfo {
        String apiPath, methodName, isDeprecated, controllerName, repoPath; String[] git1, git2, git3;
        String getApiPath() { return apiPath; }
    }
}