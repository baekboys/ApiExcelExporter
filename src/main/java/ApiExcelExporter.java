import com.github.javaparser.StaticJavaParser;
import com.github.javaparser.ast.CompilationUnit;
import com.github.javaparser.ast.body.ClassOrInterfaceDeclaration;
import com.github.javaparser.ast.body.MethodDeclaration;
import com.github.javaparser.ast.expr.*;
import com.github.javaparser.ast.comments.Comment;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.nio.file.*;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.regex.*;
import java.util.stream.Collectors;

/**
 * 프로젝트명: ApiExcelExporter (Bitbucket 관리형)
 * Version: 13.4 ('관련메뉴(미사용시)' 자동 매핑 기능 추가)
 * 반영사항:
 * 1. [자동화] '관련메뉴(미사용시)' 컬럼 우선순위 로직 적용 (ApiOperation > description > h3 > 주석첫줄) [cite: 2026-03-10]
 * 2. [로그 유지] parallelStream 내부 로그 형식 및 API 발견 상세 로그 (└ [Found]) 유지 [cite: 2026-03-10]
 * 3. [기능 유지] REST API 지능형 분석: 가변인자 및 단순 액션 제외 핵심 리소스명 추출 [cite: 2026-03-10]
 * 4. [주석 유지] 소스 코드 내 모든 상세 로직 주석 및 Javadoc 완벽 보존 [cite: 2026-02-05]
 * 5. [성능] i9-13900 환경 최적화 parallelStream 분석 및 대량 데이터 안정성 확보 [cite: 2026-02-23]
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

    /** [v12.0 신규] 관리용 팀 명칭 */
    private static String TEAM_NAME = "";

    /** [v12.0 신규] 관리용 담당자 명칭 */
    private static String MANAGER_NAME = "";

    /** [v11.3 신규] 미사용 의심 판별 기준 호출수 */
    private static long NOT_USE_LIMIT_COUNT = 0;

    /** [v11.3 신규] 미사용 의심 판별 기준일 (YYYY-MM-DD) */
    private static String LAST_COMMIT_DATE = "1900-01-01";

    /** 설정 파일 로드 성공 여부 */
    private static boolean isConfigLoaded = false;

    // ==========================================================================================
    // [ 2. 분석 엔진 및 로깅 전용 변수 ]
    // ==========================================================================================

    private static final List<String> MAPPING_ANNS = Arrays.asList("RequestMapping", "GetMapping", "PostMapping", "PutMapping", "DeleteMapping", "PatchMapping");
    private static final List<String> RUNTIME_LOGS = Collections.synchronizedList(new ArrayList<>());
    private static String logPath = "";
    private static final AtomicInteger PROCESSED_COUNT = new AtomicInteger(0);

    // ==========================================================================================

    public static void main(String[] args) {
        // 1. 설정값 먼저 로드 (항목별 상세 로그 출력 포함)
        loadExternalConfig();

        if (OUTPUT_DIR.isEmpty()) {
            System.err.println("[ERROR] OUTPUT_DIR이 설정되지 않았습니다. config.properties를 확인하세요.");
            return;
        }

        File dir = new File(OUTPUT_DIR);
        if (!dir.exists()) dir.mkdirs();

        long startTime = System.currentTimeMillis();
        // [v11.8] 날짜 형식 변경 (yyyy-MM-dd_추출)
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd'_추출'"));

        // 2. [로그] 실행 시작 정보 상세 기록
        System.out.println("===============================================================");
        System.out.println("[START] " + REPO_NAME + " API 추출 및 Whatap 통합 시작 (v13.4)");
        System.out.println("[INFO] 관리 정보: 팀[" + TEAM_NAME + "] / 담당자[" + MANAGER_NAME + "]");
        System.out.println("===============================================================");

        // 3. Whatap 통계 모듈 가동
        Map<String, long[]> whatapStats = WhatapApiCounter.getApiStats();
        WhatapApiCounter.generateExcelReport(timestamp);

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

            // [v13.2] i9-13900 병렬 분석 루프 및 로그 포맷 적용 [cite: 2026-03-10]
            controllerFiles.parallelStream().forEach(file -> {
                String relativePath = rootPathObj.relativize(file).toString();
                int current = PROCESSED_COUNT.incrementAndGet();

                List<String[]> gitHistories = getRecentGitHistories(relativePath, ROOT_PATH, 3);

                StringBuilder fileLog = new StringBuilder();
                fileLog.append(String.format("\n[%d/%d] 분석: %s", current, total, file.getFileName()));
                fileLog.append(String.format(" (최신커밋: %s | %s)", gitHistories.get(0)[0], gitHistories.get(0)[1]));

                // v13.3: fileLog를 전달하여 발견된 API 상세 정보를 누적합니다. [cite: 2026-03-10]
                allApiList.addAll(extractApisHybrid(file, relativePath, gitHistories, fileLog));
                System.out.print(fileLog.toString());
                synchronized (RUNTIME_LOGS) { RUNTIME_LOGS.add(fileLog.toString()); }
            });
        } catch (Exception e) { addExceptionLog("디렉토리 탐색 오류", e); return; }

        allApiList.sort(Comparator.comparing(ApiInfo::getApiPath));

        // 4. 결과 파일명 확정
        String baseFileName = String.format("API목록_(%s)_(컨트롤러  %d개 & API %d개)_(%s)",
                REPO_NAME, totalFiles, allApiList.size(), timestamp);

        logPath = OUTPUT_DIR + File.separator + baseFileName + ".log";
        saveInitialLogsToPath();

        File finalExcelFile = new File(OUTPUT_DIR, baseFileName + ".xlsx");

        // 5. 엑셀 리포트 생성
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(finalExcelFile)) {

            String sheetName = "API분석_" + REPO_NAME;
            if (sheetName.length() > 31) sheetName = sheetName.substring(0, 31);
            Sheet sheet = workbook.createSheet(sheetName);
            CreationHelper helper = workbook.getCreationHelper();

            // --- 셀 스타일 정의 ---
            CellStyle greyH = createStyle(workbook, IndexedColors.GREY_25_PERCENT.getIndex(), true, true);
            CellStyle yellowH = createStyle(workbook, IndexedColors.YELLOW.getIndex(), true, true);
            CellStyle orangeH = createStyle(workbook, IndexedColors.ORANGE.getIndex(), true, true);
            CellStyle blueH = createStyle(workbook, IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex(), true, true);
            CellStyle ivoryH = createStyle(workbook, IndexedColors.LEMON_CHIFFON.getIndex(), true, true);

            CellStyle centerD = createStyle(workbook, null, false, true);
            CellStyle leftD = createStyle(workbook, null, false, false);
            CellStyle numD = workbook.createCellStyle(); numD.setDataFormat(workbook.createDataFormat().getFormat("#,##0"));
            numD.setBorderBottom(BorderStyle.THIN); numD.setBorderTop(BorderStyle.THIN); numD.setBorderLeft(BorderStyle.THIN); numD.setBorderRight(BorderStyle.THIN);

            CellStyle dateD = createStyle(workbook, null, false, true);
            dateD.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd"));

            // 박스 스타일 정의
            CellStyle blueH_TopLeft = createStyle(workbook, IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex(), true, true);
            blueH_TopLeft.setBorderTop(BorderStyle.THICK); blueH_TopLeft.setBorderLeft(BorderStyle.THICK);
            CellStyle blueH_Top = createStyle(workbook, IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex(), true, true);
            blueH_Top.setBorderTop(BorderStyle.THICK);
            CellStyle blueH_TopRight = createStyle(workbook, IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex(), true, true);
            blueH_TopRight.setBorderTop(BorderStyle.THICK); blueH_TopRight.setBorderRight(BorderStyle.THICK);

            CellStyle boxLeft = createStyle(workbook, null, false, true); boxLeft.setBorderLeft(BorderStyle.THICK);
            CellStyle boxRight = createStyle(workbook, null, false, true); boxRight.setBorderRight(BorderStyle.THICK);
            CellStyle boxLeftLeftAlign = createStyle(workbook, null, false, false); boxLeftLeftAlign.setBorderLeft(BorderStyle.THICK);
            CellStyle boxRightLeftAlign = createStyle(workbook, null, false, false); boxRightLeftAlign.setBorderRight(BorderStyle.THICK);

            CellStyle boxBottom = createStyle(workbook, null, false, true); boxBottom.setBorderBottom(BorderStyle.THICK);
            CellStyle boxBottomLeft = createStyle(workbook, null, false, true); boxBottomLeft.setBorderBottom(BorderStyle.THICK); boxBottomLeft.setBorderLeft(BorderStyle.THICK);
            CellStyle boxBottomRight = createStyle(workbook, null, false, true); boxBottomRight.setBorderBottom(BorderStyle.THICK); boxBottomRight.setBorderRight(BorderStyle.THICK);

            CellStyle highRiskS = createStyle(workbook, IndexedColors.ROSE.getIndex(), true, true);
            Font redF = workbook.createFont(); redF.setColor(IndexedColors.RED.getIndex()); redF.setBold(true);
            highRiskS.setFont(redF);

            CellStyle midRiskS = createStyle(workbook, IndexedColors.YELLOW.getIndex(), true, true);
            Font goldF = workbook.createFont(); goldF.setColor(IndexedColors.GOLD.getIndex()); goldF.setBold(true);
            midRiskS.setFont(goldF);

            CellStyle lowRiskS = createStyle(workbook, IndexedColors.LIGHT_GREEN.getIndex(), true, true);
            Font darkGreenF = workbook.createFont(); darkGreenF.setColor(IndexedColors.DARK_GREEN.getIndex()); darkGreenF.setBold(true);
            lowRiskS.setFont(darkGreenF);

            CellStyle linkD = createStyle(workbook, null, false, false);
            Font linkFont = workbook.createFont(); linkFont.setColor(IndexedColors.BLUE.getIndex()); linkFont.setUnderline(Font.U_SINGLE);
            linkD.setFont(linkFont);
            CellStyle depColumnStyle = createStyle(workbook, IndexedColors.GREY_25_PERCENT.getIndex(), false, true);

            sheet.createFreezePane(4, 1);

            // [v13.3] 헤더 구성 유지 [cite: 2026-03-10]
            String[] headers = {"순번","추출일자","레파지토리","API 경로","전체 URL","repository path","컨트롤러명","호출메소드",
                    "프로그램ID(자동추출)","전체주석(참고용)","description주석(참고용)","ApiOperation(참고용)","Deprecated","커밋일자1","커밋터1","코멘트1",
                    "커밋일자2","커밋터2","코멘트2","커밋일자3","커밋터3","코멘트3",
                    "호출건수(APM추출필요)","미사용 의심건","팀","담당자","미사용 검토결과","관련메뉴(미사용시)",
                    "조치예정일자","조치일자","관련티켓","조치담당자","비고"};

            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                if (i <= 4) cell.setCellStyle(greyH);
                else if (i <= 11) cell.setCellStyle(yellowH);
                else if (i <= 23) cell.setCellStyle(orangeH);
                else if (i >= 24 && i <= 27) {
                    if (i == 24) cell.setCellStyle(blueH_TopLeft);
                    else if (i == 27) cell.setCellStyle(blueH_TopRight);
                    else cell.setCellStyle(blueH_Top);
                }
                else cell.setCellStyle(ivoryH);
            }
            sheet.setAutoFilter(new CellRangeAddress(0, 0, 0, headers.length - 1));

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            String[] reviewOptions = {"O(미사용)", "△(판단불가)", "X(사용)"};
            DataValidationConstraint constraint = validationHelper.createExplicitListConstraint(reviewOptions);
            CellRangeAddressList addressList = new CellRangeAddressList(1, Math.max(1, allApiList.size() + 1000), 26, 26);
            DataValidation validation = validationHelper.createValidation(constraint, addressList);
            validation.setSuppressDropDownArrow(true);
            validation.setShowErrorBox(true);
            sheet.addValidationData(validation);

            LocalDate now = LocalDate.now();

            for (int i = 0; i < allApiList.size(); i++) {
                ApiInfo info = allApiList.get(i);
                Row row = sheet.createRow(i + 1);
                boolean isDep = "Y".equals(info.isDeprecated);
                String fullUrl = DOMAIN + info.apiPath;

                long totalCalls = 0;
                long[] rowStats = whatapStats.get(info.apiPath);
                if (rowStats != null) for (long count : rowStats) totalCalls += count;

                String suspicionScore = "";
                int caseType = 0;
                LocalDate thresholdDate = LocalDate.parse(LAST_COMMIT_DATE);
                LocalDate latestCommitDate = getLatestDate(info.git1[0], info.git2[0], info.git3[0]);

                if (isDep && totalCalls == 0) { suspicionScore = "★★★"; caseType = 1; }
                else if (totalCalls <= NOT_USE_LIMIT_COUNT) {
                    if (latestCommitDate != null && latestCommitDate.isBefore(thresholdDate)) { suspicionScore = "★★☆"; caseType = 2; }
                    else { suspicionScore = "★☆☆"; caseType = 3; }
                }

                String autoProgId = autoExtractProgramId(info.apiPath);

                // [v13.4] 관련메뉴 자동 매핑 로직 적용 [cite: 2026-03-10]
                String autoRelatedMenu = autoPopulateRelatedMenu(info);

                String[] data = {String.valueOf(i + 1), "", REPO_NAME, info.apiPath, fullUrl, info.repoPath,
                        info.controllerName, info.methodName, autoProgId, info.fullComment, info.descriptionTag, info.apiOperationValue,
                        info.isDeprecated, info.git1[0], info.git1[1], info.git1[2], info.git2[0], info.git2[1], info.git2[2],
                        info.git3[0], info.git3[1], info.git3[2], String.valueOf(totalCalls), suspicionScore,
                        TEAM_NAME, MANAGER_NAME, "", autoRelatedMenu, "", "", "", "", ""};

                boolean isLastRow = (i == allApiList.size() - 1);

                for (int j = 0; j < data.length; j++) {
                    Cell cell = row.createCell(j);
                    if (j == 1) {
                        cell.setCellValue(now);
                        cell.setCellStyle(dateD);
                    } else if (j == 22) {
                        cell.setCellValue(totalCalls);
                        if (totalCalls >= 0 && totalCalls <= NOT_USE_LIMIT_COUNT) cell.setCellStyle(highRiskS);
                        else cell.setCellStyle(numD);
                    } else if (j == 23) {
                        cell.setCellValue(suspicionScore);
                        if (caseType == 1) cell.setCellStyle(highRiskS);
                        else if (caseType == 2) cell.setCellStyle(midRiskS);
                        else if (caseType == 3) cell.setCellStyle(lowRiskS);
                        else cell.setCellStyle(centerD);
                    } else {
                        cell.setCellValue(data[j]);
                        // 주석 3종 왼쪽 정렬 적용 [cite: 2026-03-10]
                        boolean isCenter = (j==0 || j==2 || (j>=6 && j<=8) || (j>=12 && j<=21) || (j>=24));

                        if (j == 12 && isDep) cell.setCellStyle(depColumnStyle);
                        else if (j == 4) {
                            cell.setCellStyle(linkD);
                            try {
                                String encodedUrl = fullUrl.replace("{", "%7B").replace("}", "%7D");
                                Hyperlink link = helper.createHyperlink(HyperlinkType.URL);
                                link.setAddress(encodedUrl);
                                cell.setHyperlink(link);
                            } catch (Exception ignored) { }
                        } else {
                            if (j >= 24 && j <= 27) {
                                if (isLastRow) {
                                    if (j == 24) cell.setCellStyle(boxBottomLeft);
                                    else if (j == 27) cell.setCellStyle(boxBottomRight);
                                    else cell.setCellStyle(boxBottom);
                                } else {
                                    if (j == 24) cell.setCellStyle(isCenter ? boxLeft : boxLeftLeftAlign);
                                    else if (j == 27) cell.setCellStyle(isCenter ? boxRight : boxRightLeftAlign);
                                    else cell.setCellStyle(isCenter ? centerD : leftD);
                                }
                            } else {
                                cell.setCellStyle(isCenter ? centerD : leftD);
                            }
                        }
                    }
                }
            }

            sheet.setColumnWidth(1, 4000);
            sheet.setColumnWidth(3, 14500); sheet.setColumnWidth(4, 8500);
            sheet.setColumnWidth(5, 11500); sheet.setColumnWidth(6, 5500); sheet.setColumnWidth(7, 5500);
            sheet.setColumnWidth(8, 5500);
            sheet.setColumnWidth(9, 5800);
            sheet.setColumnWidth(10, 5800);
            sheet.setColumnWidth(11, 5800);
            sheet.setColumnWidth(24, 4000);
            sheet.setColumnWidth(27, 6000);
            for (int i = 13; i < headers.length; i++) if(i<24 || i>27) sheet.setColumnWidth(i, 4200);

            workbook.write(fos);
            addLog("\n[SUCCESS] 통합 엑셀 저장 완료: " + finalExcelFile.getName());
        } catch (Exception e) { addExceptionLog("엑셀 저장 중 오류", e); }

        addLog("\n[FINISH] 전체 분석 작업 종료: " + (System.currentTimeMillis() - startTime) / 1000 + "초 소요");
    }

    /** [v13.4] 관련메뉴(미사용시) 자동 매핑 헬퍼 로직 [cite: 2026-03-10]
     * 우선순위: ApiOperation > description > Javadoc h3 태그 > 주석 첫 문장 */
    private static String autoPopulateRelatedMenu(ApiInfo info) {
        // 1. ApiOperation 우선순위 [cite: 2026-03-10]
        if (info.apiOperationValue != null && !"-".equals(info.apiOperationValue) && !info.apiOperationValue.trim().isEmpty()) {
            return info.apiOperationValue;
        }
        // 2. description 주석 우선순위 [cite: 2026-03-10]
        if (info.descriptionTag != null && !"-".equals(info.descriptionTag) && !info.descriptionTag.trim().isEmpty()) {
            return info.descriptionTag;
        }
        // 3. Javadoc 분석 (h3 태그 및 첫 문장) [cite: 2026-03-10]
        if (info.fullComment != null && !"-".equals(info.fullComment)) {
            // h3 태그 추출 시도 [cite: 2026-03-10]
            Pattern h3Pattern = Pattern.compile("<h3>(.*?)</h3>", Pattern.CASE_INSENSITIVE);
            Matcher matcher = h3Pattern.matcher(info.fullComment);
            if (matcher.find()) return matcher.group(1).trim();

            // 첫 문장/줄 추출 (구분자 기준) [cite: 2026-03-10]
            String comment = info.fullComment.trim();
            if (!comment.isEmpty()) {
                String firstPart = comment.split("[.!?:]")[0];
                return firstPart.trim().length() > 2 ? firstPart.trim() : comment;
            }
        }
        return "-";
    }

    /** [v13.1] REST API 지능형 핵심 리소스명 추출 로직 유지 [cite: 2026-03-10] */
    private static String autoExtractProgramId(String path) {
        if (path == null || path.isEmpty() || "/".equals(path)) return "-";
        if (path.contains(".")) {
            int lastSlash = path.lastIndexOf("/");
            String filePart = (lastSlash != -1) ? path.substring(lastSlash + 1) : path;
            int dotIdx = filePart.lastIndexOf(".");
            String nameOnly = (dotIdx != -1) ? filePart.substring(0, dotIdx) : filePart;
            int underIdx = nameOnly.lastIndexOf("_");
            return (underIdx != -1) ? nameOnly.substring(0, underIdx) : nameOnly;
        }
        String[] segments = path.split("/");
        List<String> validNouns = new ArrayList<>();
        List<String> actions = Arrays.asList("new", "edit", "update", "delete", "create", "list", "save", "view");
        for (String s : segments) {
            if (!s.isEmpty() && !s.startsWith("{") && !actions.contains(s.toLowerCase())) {
                validNouns.add(s);
            }
        }
        return validNouns.isEmpty() ? "-" : validNouns.get(validNouns.size() - 1);
    }

    private static LocalDate getLatestDate(String d1, String d2, String d3) {
        List<LocalDate> dates = new ArrayList<>();
        try { if(!"-".equals(d1)) dates.add(LocalDate.parse(d1)); } catch(Exception ignored){}
        try { if(!"-".equals(d2)) dates.add(LocalDate.parse(d2)); } catch(Exception ignored){}
        try { if(!"-".equals(d3)) dates.add(LocalDate.parse(d3)); } catch(Exception ignored){}
        return dates.stream().max(Comparator.naturalOrder()).orElse(null);
    }

    private static void loadExternalConfig() {
        Properties prop = new Properties();
        File configFile = new File("config.properties");
        if (configFile.exists()) {
            try (InputStreamReader isr = new InputStreamReader(new FileInputStream(configFile), StandardCharsets.UTF_8)) {
                prop.load(isr);
                if (prop.getProperty("REPO_NAME") != null) REPO_NAME = prop.getProperty("REPO_NAME").trim();
                if (prop.getProperty("DOMAIN") != null) DOMAIN = prop.getProperty("DOMAIN").trim();
                if (prop.getProperty("ROOT_PATH") != null) ROOT_PATH = prop.getProperty("ROOT_PATH").trim();
                if (prop.getProperty("OUTPUT_DIR") != null) OUTPUT_DIR = prop.getProperty("OUTPUT_DIR").trim();
                if (prop.getProperty("GIT_BIN_PATH") != null) GIT_BIN_PATH = prop.getProperty("GIT_BIN_PATH").trim();
                TEAM_NAME = prop.getProperty("TEAM_NAME", "").trim();
                MANAGER_NAME = prop.getProperty("MANAGER_NAME", "").trim();
                if (prop.getProperty("NOT_USE_LIMIT_COUNT") != null) NOT_USE_LIMIT_COUNT = Long.parseLong(prop.getProperty("NOT_USE_LIMIT_COUNT").trim());
                if (prop.getProperty("LAST_COMMIT_DATE") != null) LAST_COMMIT_DATE = prop.getProperty("LAST_COMMIT_DATE").trim();
                isConfigLoaded = true;
            } catch (IOException e) { System.err.println("[ERROR] 설정 로드 중 오류: " + e.getMessage()); }
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
        CompilationUnit cu = StaticJavaParser.parse(new String(Files.readAllBytes(filePath), StandardCharsets.UTF_8));
        String classPath = "";
        Optional<ClassOrInterfaceDeclaration> mainClass = cu.findFirst(ClassOrInterfaceDeclaration.class);
        if (mainClass.isPresent()) {
            Optional<AnnotationExpr> classAnn = mainClass.get().getAnnotationByName("RequestMapping");
            if (classAnn.isPresent()) { List<String> cpList = getPathsFromAnn(classAnn.get()); if (!cpList.isEmpty()) classPath = cpList.get(0).trim(); }
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
                        info.methodName = method.getNameAsString(); info.isDeprecated = method.isAnnotationPresent("Deprecated") ? "Y" : "N";
                        info.controllerName = filePath.getFileName().toString(); info.repoPath = (REPO_NAME + "/" + relPath).replace("\\", "/");
                        info.git1 = git.get(0); info.git2 = git.get(1); info.git3 = git.get(2);

                        if (method.getComment().isPresent()) {
                            String full = method.getComment().get().getContent();
                            info.fullComment = full.replaceAll("\\r|\\n|\\*", " ").trim();
                            Matcher dM = Pattern.compile("@description\\s*:\\s*([^\\n\\r*]+)").matcher(full);
                            info.descriptionTag = dM.find() ? dM.group(1).trim() : "-";
                        } else {
                            info.fullComment = "-"; info.descriptionTag = "-";
                        }
                        info.apiOperationValue = extractAnnotationValue(method, "ApiOperation", "value");
                        apis.add(info);
                        log.append("\n    └ [Found] ").append(info.apiPath);
                    }
                }
            }
        }
        return apis;
    }

    private static String extractAnnotationValue(MethodDeclaration method, String annName, String attrName) {
        Optional<AnnotationExpr> ann = method.getAnnotationByName(annName);
        if (ann.isPresent() && ann.get() instanceof NormalAnnotationExpr) {
            return ((NormalAnnotationExpr) ann.get()).getPairs().stream()
                    .filter(p -> p.getNameAsString().equals(attrName))
                    .map(p -> p.getValue().toString().replaceAll("\"", ""))
                    .findFirst().orElse("-");
        } else if (ann.isPresent() && ann.get() instanceof SingleMemberAnnotationExpr) {
            return ((SingleMemberAnnotationExpr) ann.get()).getMemberValue().toString().replaceAll("\"", "");
        }
        return "-";
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
                int searchLimit = Math.min(mMatcher.end() + 1000, clean.length());
                Matcher mName = Pattern.compile("(?:public|private|protected)\\s+[\\w<>,\\s]+\\s+(\\w+)\\s*\\(").matcher(clean.substring(mMatcher.end(), searchLimit));
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
                            info.methodName = mName.group(1); info.isDeprecated = clean.substring(Math.max(0, mMatcher.start() - 300), mMatcher.start()).contains("@Deprecated") ? "Y" : "N";
                            info.controllerName = filePath.getFileName().toString(); info.repoPath = (REPO_NAME + "/" + relPath).replace("\\", "/");
                            info.git1 = git.get(0); info.git2 = git.get(1); info.git3 = git.get(2);

                            String headArea = raw.substring(Math.max(0, mMatcher.start() - 1000), mMatcher.start());
                            Matcher cM = Pattern.compile("/\\*\\*(.*?)\\*/", Pattern.DOTALL).matcher(headArea);
                            if (cM.find()) {
                                String full = cM.group(1);
                                info.fullComment = full.replaceAll("\\r|\\n|\\*", " ").trim();
                                Matcher dM = Pattern.compile("@description\\s*:\\s*([^\\n\\r*]+)").matcher(full);
                                info.descriptionTag = dM.find() ? dM.group(1).trim() : "-";
                            } else {
                                info.fullComment = "-"; info.descriptionTag = "-";
                            }
                            Matcher aM = Pattern.compile("@ApiOperation\\s*\\(\\s*value\\s*=\\s*\"([^\"]+)\"").matcher(headArea);
                            info.apiOperationValue = aM.find() ? aM.group(1) : "-";
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
        if (ann instanceof NormalAnnotationExpr) { value = ((NormalAnnotationExpr) ann).getPairs().stream().filter(p -> p.getNameAsString().equals("value") || p.getNameAsString().equals("path")).map(MemberValuePair::getValue).findFirst().orElse(null); }
        if (value instanceof ArrayInitializerExpr) { for (Expression expr : ((ArrayInitializerExpr) value).getValues()) if (expr instanceof StringLiteralExpr) paths.add(((StringLiteralExpr) expr).getValue()); }
        else if (value instanceof StringLiteralExpr) paths.add(((StringLiteralExpr) value).getValue());
        return paths;
    }

    private static void addLog(String msg) { System.out.println(msg); if (!logPath.isEmpty()) { try (FileWriter fw = new FileWriter(logPath, true); PrintWriter pw = new PrintWriter(fw)) { pw.println(msg); } catch (IOException ignored) {} } }
    private static void saveInitialLogsToPath() { try (FileWriter fw = new FileWriter(logPath, false); PrintWriter pw = new PrintWriter(fw)) { pw.println("==============================================================="); pw.println("[START] " + REPO_NAME + " API 추출 및 Whatap 통합 시작 (v13.4)"); pw.println("==============================================================="); synchronized (RUNTIME_LOGS) { for (String l : RUNTIME_LOGS) pw.println(l); } } catch (IOException ignored) {} }
    private static void addExceptionLog(String title, Exception e) { StringWriter sw = new StringWriter(); e.printStackTrace(new PrintWriter(sw)); addLog("\n[ERROR] " + title + "\n" + sw.toString()); }

    private static List<String[]> getRecentGitHistories(String rel, String root, int c) {
        List<String[]> h = new ArrayList<>(); for (int i = 0; i < c; i++) h.add(new String[]{"-", "-", "No History"});
        try {
            Process p = new ProcessBuilder(GIT_BIN_PATH, "log", "-" + c, "--pretty=format:%as|%an|%s", "--", rel).directory(new File(root)).start();
            try (BufferedReader r = new BufferedReader(new InputStreamReader(p.getInputStream(), StandardCharsets.UTF_8))) {
                List<String> lines = new ArrayList<>(); String l; while ((l = r.readLine()) != null) lines.add(l); Collections.reverse(lines);
                for (int i = 0; i < lines.size(); i++) { String[] parts = lines.get(i).split("\\|", 3); h.set(i, new String[]{parts[0], parts[1], parts.length > 2 ? parts[2] : ""}); }
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
        String apiPath, methodName, isDeprecated, controllerName, repoPath;
        String descriptionTag, apiOperationValue, fullComment;
        String[] git1, git2, git3; String getApiPath() { return apiPath; }
    }
}