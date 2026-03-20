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
 * Version: 13.14 (메소드 단위 상세 분석 추적 로깅 탑재)
 * 반영사항:
 * 1. [로깅 보완] JavaParser 및 Regex 분석 시 모든 메소드를 로깅하고, 스킵된 사유([Skip])를 상세히 기록 [cite: 2026-03-20]
 * 2. [기능 유지] config.properties의 API_PATH_PREFIX 값을 읽어 모든 추출된 API 경로 앞에 일괄 추가 [cite: 2026-03-12]
 * 3. [기능 유지] PATH_CONSTANTS를 이용한 상수 치환(evaluateExpression) 및 텍스트 정제(cleanMeaningfulText) 완벽 보존 [cite: 2026-03-12]
 * 4. [레이아웃 유지] 대량 데이터 조건부 서식 동적 범위(4,000건 이상), 관련메뉴 위계 로직 완벽 보존 [cite: 2026-03-11]
 * 5. [성능/유지] i9-13900 병렬 분석, 기존 로그 포맷([Found]), 소스 코드 내 모든 상세 주석 완벽 보존 [cite: 2026-02-05, 2026-02-23]
 */
public class ApiExcelExporter {

    // ==========================================================================================
    // [ 1. 내부 기본 설정부 ] - config.properties를 반드시 작성하세요.
    // ==========================================================================================

    /** [핵심변수 1] 레파지토리 이름 : 파일명 생성 시 식별자로 활용 */
    private static String REPO_NAME = "";

    /** [핵심변수 2] 기본 도메인 주소 : 엑셀 내 전체 URL 하이퍼링크 생성용 */
    private static String DOMAIN = "";

    /** [핵심변수 3] 분석할 Java 소스 로컬 절대 경로 */
    private static String ROOT_PATH = "";

    /** [핵심변수 4] 결과 저장 디렉토리 물리적 경로 */
    private static String OUTPUT_DIR = "";

    /** [핵심변수 5] Git 실행 경로 : 환경변수 미등록 PC 대응 */
    private static String GIT_BIN_PATH = "git";

    /** [v12.0 신규] 관리용 팀 명칭 */
    private static String TEAM_NAME = "";

    /** [v12.0 신규] 관리용 담당자 명칭 */
    private static String MANAGER_NAME = "";

    /** [v11.3 신규] 미사용 의심 판별 기준 호출수 */
    private static long NOT_USE_LIMIT_COUNT = 0;

    /** [v11.3 신규] 미사용 의심 판별 기준일 (YYYY-MM-DD) */
    private static String LAST_COMMIT_DATE = "1900-01-01";

    /** [v13.6 신규] Whatap 연동 여부 : N일 경우 호출건수 등을 표시하지 않음 [cite: 2026-03-11] */
    private static String WHATAP_ENABLED = "Y";

    /** [v13.12 신규] API 경로 내 상수 치환용 맵 (config.properties의 PATH_CONSTANTS) [cite: 2026-03-12] */
    private static final Map<String, String> PATH_CONSTANTS_MAP = new HashMap<>();

    /** [v13.13 신규] 전체 API 경로 앞에 일괄 추가할 Prefix (config.properties의 API_PATH_PREFIX) [cite: 2026-03-12] */
    private static String API_PATH_PREFIX = "";

    /** 설정 파일 로드 성공 여부 플래그 */
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
        loadExternalConfig();

        if (OUTPUT_DIR.isEmpty()) {
            System.err.println("[ERROR] OUTPUT_DIR이 설정되지 않았습니다. config.properties를 확인하세요.");
            return;
        }

        File dir = new File(OUTPUT_DIR);
        if (!dir.exists()) dir.mkdirs();

        long startTime = System.currentTimeMillis();
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd'_추출'"));

        System.out.println("===============================================================");
        System.out.println("[START] " + REPO_NAME + " API 추출 및 Whatap 통합 시작 (v13.14)");
        System.out.println("[INFO] 관리 정보: 팀[" + TEAM_NAME + "] / 담당자[" + MANAGER_NAME + "]");
        System.out.println("===============================================================");

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
        } catch (Exception e) { addExceptionLog("디렉토리 탐색 오류", e); return; }

        allApiList.sort(Comparator.comparing(ApiInfo::getApiPath));

        String baseFileName = String.format("API목록_(%s)_(컨트롤러  %d개 & API %d개)_(%s)",
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

            // --- 스타일 정의 부 ---
            CellStyle greyH = createStyle(workbook, IndexedColors.GREY_25_PERCENT.getIndex(), true, true);
            CellStyle yellowH = createStyle(workbook, IndexedColors.YELLOW.getIndex(), true, true);
            CellStyle orangeH = createStyle(workbook, IndexedColors.ORANGE.getIndex(), true, true);
            CellStyle blueH = createStyle(workbook, IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex(), true, true);
            CellStyle ivoryH = createStyle(workbook, IndexedColors.LEMON_CHIFFON.getIndex(), true, true);
            CellStyle leftD = createStyle(workbook, null, false, false);
            CellStyle centerD = createStyle(workbook, null, false, true);
            CellStyle numD = workbook.createCellStyle(); numD.setDataFormat(workbook.createDataFormat().getFormat("#,##0"));
            numD.setBorderBottom(BorderStyle.THIN); numD.setBorderTop(BorderStyle.THIN); numD.setBorderLeft(BorderStyle.THIN); numD.setBorderRight(BorderStyle.THIN);
            CellStyle dateD = createStyle(workbook, null, false, true); dateD.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd"));
            CellStyle depColumnStyle = createStyle(workbook, IndexedColors.GREY_25_PERCENT.getIndex(), false, true);
            CellStyle linkD = createStyle(workbook, null, false, false); Font linkFont = workbook.createFont(); linkFont.setColor(IndexedColors.BLUE.getIndex()); linkFont.setUnderline(Font.U_SINGLE); linkD.setFont(linkFont);

            CellStyle boxLeft = createStyle(workbook, null, false, true); boxLeft.setBorderLeft(BorderStyle.THICK);
            CellStyle boxRight = createStyle(workbook, null, false, true); boxRight.setBorderRight(BorderStyle.THICK);
            CellStyle boxLeftLeftAlign = createStyle(workbook, null, false, false); boxLeftLeftAlign.setBorderLeft(BorderStyle.THICK);
            CellStyle boxRightLeftAlign = createStyle(workbook, null, false, false); boxRightLeftAlign.setBorderRight(BorderStyle.THICK);
            CellStyle boxBottom = createStyle(workbook, null, false, true); boxBottom.setBorderBottom(BorderStyle.THICK);
            CellStyle boxBottomLeft = createStyle(workbook, null, false, true); boxBottomLeft.setBorderBottom(BorderStyle.THICK); boxBottomLeft.setBorderLeft(BorderStyle.THICK);
            CellStyle boxBottomRight = createStyle(workbook, null, false, true); boxBottomRight.setBorderBottom(BorderStyle.THICK); boxBottomRight.setBorderRight(BorderStyle.THICK);

            sheet.createFreezePane(4, 1);

            String[] headers = {"순번","추출일자","레파지토리","API 경로","전체 URL","repository path","컨트롤러명","호출메소드",
                    "프로그램ID(자동추출)","ApiOperation(참고용)","description주석(참고용)","메소드주석(참고용)",
                    "RequestProperty(참고용)","컨트롤러RequestProperty(참고용)","컨트롤러주석(참고용)","Deprecated",
                    "커밋일자1","커밋터1","코멘트1","커밋일자2","커밋터2","코멘트2","커밋일자3","커밋터3","코멘트3",
                    "호출건수(APM추출필요)","미사용 의심건","팀","담당자","미사용 검토결과","관련메뉴(미사용시)",
                    "조치예정일자","조치일자","관련티켓","조치담당자","비고"};

            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                if (i <= 4) cell.setCellStyle(greyH);
                else if (i <= 14) cell.setCellStyle(yellowH);
                else if (i <= 26) cell.setCellStyle(orangeH);
                else if (i >= 27 && i <= 30) {
                    CellStyle style = createStyle(workbook, IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex(), true, true);
                    style.setBorderTop(BorderStyle.THICK); if (i == 27) style.setBorderLeft(BorderStyle.THICK); if (i == 30) style.setBorderRight(BorderStyle.THICK);
                    cell.setCellStyle(style);
                }
                else cell.setCellStyle(ivoryH);
            }
            sheet.setAutoFilter(new CellRangeAddress(0, 0, 0, headers.length - 1));

            DataValidationHelper validationHelper = sheet.getDataValidationHelper();
            int maxRowIndex = Math.max(5000, allApiList.size() + 1000);
            String lastRowStr = String.valueOf(maxRowIndex + 1);

            CellRangeAddressList suspicionAddressList = new CellRangeAddressList(1, maxRowIndex, 26, 26);
            DataValidation suspicionValidation = validationHelper.createValidation(validationHelper.createExplicitListConstraint(new String[]{"★☆☆", "★★☆", "★★★"}), suspicionAddressList);
            sheet.addValidationData(suspicionValidation);

            CellRangeAddressList addressList = new CellRangeAddressList(1, maxRowIndex, 29, 29);
            DataValidation validation = validationHelper.createValidation(validationHelper.createExplicitListConstraint(new String[]{"O(미사용)", "△(판단불가)", "X(사용)"}), addressList);
            sheet.addValidationData(validation);

            SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();
            String callCountFormula = String.format("AND(Z2<>\"\", Z2<=%d)", NOT_USE_LIMIT_COUNT);
            ConditionalFormattingRule callCountRule = sheetCF.createConditionalFormattingRule(callCountFormula);
            PatternFormatting callCountFill = callCountRule.createPatternFormatting();
            callCountFill.setFillBackgroundColor(IndexedColors.ROSE.getIndex()); callCountFill.setFillPattern(PatternFormatting.SOLID_FOREGROUND);
            sheetCF.addConditionalFormatting(new CellRangeAddress[]{CellRangeAddress.valueOf("Z2:Z" + lastRowStr)}, callCountRule);

            ConditionalFormattingRule rule3 = sheetCF.createConditionalFormattingRule(ComparisonOperator.EQUAL, "\"★★★\"");
            rule3.createPatternFormatting().setFillBackgroundColor(IndexedColors.ROSE.getIndex()); rule3.createPatternFormatting().setFillPattern(PatternFormatting.SOLID_FOREGROUND);
            ConditionalFormattingRule rule2 = sheetCF.createConditionalFormattingRule(ComparisonOperator.EQUAL, "\"★★☆\"");
            rule2.createPatternFormatting().setFillBackgroundColor(IndexedColors.YELLOW.getIndex()); rule2.createPatternFormatting().setFillPattern(PatternFormatting.SOLID_FOREGROUND);
            ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule(ComparisonOperator.EQUAL, "\"★☆☆\"");
            rule1.createPatternFormatting().setFillBackgroundColor(IndexedColors.LIGHT_GREEN.getIndex()); rule1.createPatternFormatting().setFillPattern(PatternFormatting.SOLID_FOREGROUND);
            sheetCF.addConditionalFormatting(new CellRangeAddress[]{CellRangeAddress.valueOf("AA2:AA" + lastRowStr)}, new ConditionalFormattingRule[]{rule3, rule2, rule1});

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
                LocalDate thresholdDate = LocalDate.parse(LAST_COMMIT_DATE);
                LocalDate latestCommitDate = getLatestDate(info.git1[0], info.git2[0], info.git3[0]);
                if (isDep && totalCalls == 0) suspicionScore = "★★★";
                else if (totalCalls <= NOT_USE_LIMIT_COUNT) {
                    if (latestCommitDate != null && latestCommitDate.isBefore(thresholdDate)) suspicionScore = "★★☆";
                    else suspicionScore = "★☆☆";
                }

                String autoRelatedMenu = autoPopulateRelatedMenu(info);
                boolean isWhatapOn = "Y".equalsIgnoreCase(WHATAP_ENABLED);

                String[] data = {String.valueOf(i + 1), "", REPO_NAME, info.apiPath, fullUrl, info.repoPath,
                        info.controllerName, info.methodName, autoExtractProgramId(info.apiPath),
                        info.apiOperationValue, info.descriptionTag, info.fullComment,
                        info.requestPropertyValue, info.controllerRequestPropertyValue, info.controllerComment,
                        info.isDeprecated, info.git1[0], info.git1[1], info.git1[2], info.git2[0], info.git2[1], info.git2[2],
                        info.git3[0], info.git3[1], info.git3[2], isWhatapOn ? String.valueOf(totalCalls) : "", isWhatapOn ? suspicionScore : "",
                        TEAM_NAME, MANAGER_NAME, "", autoRelatedMenu, "", "", "", "", ""};

                boolean isLastRow = (i == allApiList.size() - 1);
                for (int j = 0; j < data.length; j++) {
                    Cell cell = row.createCell(j);
                    if (j == 1) { cell.setCellValue(now); cell.setCellStyle(dateD); }
                    else if (j == 25) { if (isWhatapOn) { cell.setCellValue(totalCalls); cell.setCellStyle(numD); } else cell.setCellStyle(centerD); }
                    else if (j == 26) { cell.setCellValue(data[j]); cell.setCellStyle(centerD); }
                    else {
                        cell.setCellValue(data[j]);
                        boolean isCenter = (j==0 || j==1 || j==2 || (j>=6 && j<=8) || (j>=15 && j<=25) || (j>=27));
                        if (j == 15 && isDep) cell.setCellStyle(depColumnStyle);
                        else if (j == 4) {
                            cell.setCellStyle(linkD);
                            try { cell.setHyperlink(helper.createHyperlink(HyperlinkType.URL)); cell.getHyperlink().setAddress(fullUrl.replace("{", "%7B").replace("}", "%7D")); } catch (Exception ignored) {}
                        } else {
                            if (j >= 27 && j <= 30) {
                                if (isLastRow) { cell.setCellStyle(j==27 ? boxBottomLeft : (j==30 ? boxBottomRight : boxBottom)); }
                                else { cell.setCellStyle(j==27 ? (isCenter ? boxLeft : boxLeftLeftAlign) : (j==30 ? (isCenter ? boxRight : boxRightLeftAlign) : (isCenter ? centerD : leftD))); }
                            } else cell.setCellStyle(isCenter ? centerD : leftD);
                        }
                    }
                }
            }
            sheet.setColumnWidth(1, 4000); sheet.setColumnWidth(3, 14500); sheet.setColumnWidth(4, 8500);
            for (int k = 9; k <= 14; k++) sheet.setColumnWidth(k, 5800);
            sheet.setColumnWidth(29, 3500); sheet.setColumnWidth(30, 6000);
            workbook.write(fos);
        } catch (Exception e) { addExceptionLog("엑셀 저장 중 오류", e); }
        addLog("\n[FINISH] 작업 종료: " + (System.currentTimeMillis() - startTime) / 1000 + "초 소요");
    }

    private static String autoPopulateRelatedMenu(ApiInfo info) {
        if (info.apiOperationValue != null && !"-".equals(info.apiOperationValue) && !info.apiOperationValue.trim().isEmpty()) return info.apiOperationValue;
        if (info.descriptionTag != null && !"-".equals(info.descriptionTag) && !info.descriptionTag.trim().isEmpty()) return cleanMeaningfulText(info.descriptionTag);

        String mainCmt = cleanMeaningfulText(info.fullComment);
        if (!"-".equals(mainCmt)) return mainCmt;

        if (info.requestPropertyValue != null && !"-".equals(info.requestPropertyValue) && !info.requestPropertyValue.trim().isEmpty()) return info.requestPropertyValue;
        if (info.controllerRequestPropertyValue != null && !"-".equals(info.controllerRequestPropertyValue) && !info.controllerRequestPropertyValue.trim().isEmpty()) return info.controllerRequestPropertyValue;

        return cleanMeaningfulText(info.controllerComment);
    }

    private static String cleanMeaningfulText(String input) {
        if (input == null || "-".equals(input) || input.trim().isEmpty()) return "-";
        String cleaned = input.split("@")[0];
        cleaned = cleaned.replaceAll("<[^>]*>", " ");
        cleaned = cleaned.replaceAll("(?i)(ModelAndView|HttpServletRequest|HttpServletResponse|@ResponseBody|@RequestBody|@PathVariable)", "");
        String[] parts = cleaned.trim().split("[\\n\\r.,;]");
        for (String p : parts) {
            String trimmed = p.trim();
            if (trimmed.length() > 2) return trimmed;
        }
        return cleaned.trim().isEmpty() ? "-" : cleaned.trim();
    }

    private static String autoExtractProgramId(String path) {
        if (path == null || path.isEmpty() || "/".equals(path)) return "-";
        if (path.contains(".")) {
            String nameOnly = path.substring(path.lastIndexOf("/") + 1).split("\\.")[0];
            return nameOnly.contains("_") ? nameOnly.substring(0, nameOnly.lastIndexOf("_")) : nameOnly;
        }
        String[] segments = path.split("/");
        List<String> valid = new ArrayList<>();
        List<String> actions = Arrays.asList("new", "edit", "update", "delete", "create", "list", "save", "view");
        for (String s : segments) if (!s.isEmpty() && !s.startsWith("{") && !actions.contains(s.toLowerCase())) valid.add(s);
        return valid.isEmpty() ? "-" : valid.get(valid.size() - 1);
    }

    private static LocalDate getLatestDate(String d1, String d2, String d3) {
        List<LocalDate> dates = new ArrayList<>();
        try { if(d1!=null && !"-".equals(d1)) dates.add(LocalDate.parse(d1)); } catch(Exception ignored){}
        try { if(d2!=null && !"-".equals(d2)) dates.add(LocalDate.parse(d2)); } catch(Exception ignored){}
        try { if(d3!=null && !"-".equals(d3)) dates.add(LocalDate.parse(d3)); } catch(Exception ignored){}
        return dates.stream().max(Comparator.naturalOrder()).orElse(null);
    }

    private static void loadExternalConfig() {
        Properties prop = new Properties();
        try (InputStreamReader isr = new InputStreamReader(new FileInputStream("config.properties"), StandardCharsets.UTF_8)) {
            prop.load(isr);
            REPO_NAME = prop.getProperty("REPO_NAME", "Unknown").trim();
            DOMAIN = prop.getProperty("DOMAIN", "").trim();
            ROOT_PATH = prop.getProperty("ROOT_PATH", "").trim();
            OUTPUT_DIR = prop.getProperty("OUTPUT_DIR", "").trim();
            GIT_BIN_PATH = prop.getProperty("GIT_BIN_PATH", "git").trim();
            TEAM_NAME = prop.getProperty("TEAM_NAME", "").trim();
            MANAGER_NAME = prop.getProperty("MANAGER_NAME", "").trim();
            NOT_USE_LIMIT_COUNT = Long.parseLong(prop.getProperty("NOT_USE_LIMIT_COUNT", "0").trim());
            LAST_COMMIT_DATE = prop.getProperty("LAST_COMMIT_DATE", "1900-01-01").trim();
            WHATAP_ENABLED = prop.getProperty("WHATAP_ENABLED", "Y").trim();

            API_PATH_PREFIX = prop.getProperty("API_PATH_PREFIX", "").trim();

            String pathConstantsStr = prop.getProperty("PATH_CONSTANTS", "").trim();
            if (!pathConstantsStr.isEmpty()) {
                for (String pair : pathConstantsStr.split(",")) {
                    String[] kv = pair.split("=");
                    if (kv.length == 2) PATH_CONSTANTS_MAP.put(kv[0].trim(), kv[1].trim());
                }
            }
        } catch (Exception e) { System.err.println("[ERROR] 설정 로드 실패: " + e.getMessage()); }
    }

    private static List<ApiInfo> extractApisHybrid(Path path, String rel, List<String[]> git, StringBuilder log) {
        try { return extractWithJavaParser(path, rel, git, log); }
        catch (Exception e) { return extractWithRegex(path, rel, git, log); }
    }

    private static List<ApiInfo> extractWithJavaParser(Path filePath, String relPath, List<String[]> git, StringBuilder log) throws Exception {
        List<ApiInfo> apis = new ArrayList<>();
        CompilationUnit cu = StaticJavaParser.parse(new String(Files.readAllBytes(filePath), StandardCharsets.UTF_8));
        String classPath = ""; String controllerComment = "-"; String controllerRequestProperty = "-";
        Optional<ClassOrInterfaceDeclaration> mainClass = cu.findFirst(ClassOrInterfaceDeclaration.class);
        if (mainClass.isPresent()) {
            ClassOrInterfaceDeclaration n = mainClass.get();
            controllerComment = n.getComment().isPresent() ? n.getComment().get().getContent().replaceAll("\\r|\\n|\\*", " ").trim() : "-";
            controllerRequestProperty = extractRequestPropertyFromNode(n);
            Optional<AnnotationExpr> classAnn = n.getAnnotationByName("RequestMapping");
            if (classAnn.isPresent()) { List<String> cpList = getPathsFromAnn(classAnn.get()); if (!cpList.isEmpty()) classPath = cpList.get(0).trim(); }
        }

        for (MethodDeclaration method : cu.findAll(MethodDeclaration.class)) {
            // [v13.14 상세 로깅 추가] 분석 진입점 기록 [cite: 2026-03-20]
            log.append("\n    * [Analyze] 메소드명: ").append(method.getNameAsString());
            boolean hasMapping = false;

            for (String annName : MAPPING_ANNS) {
                Optional<AnnotationExpr> methodAnn = method.getAnnotationByName(annName);
                if (methodAnn.isPresent()) {
                    hasMapping = true;
                    List<String> subPaths = getPathsFromAnn(methodAnn.get());
                    if (subPaths.isEmpty()) {
                        subPaths.add("");
                        log.append("\n      - [Info] 매핑값 없음, 기본(\"\") 경로로 처리");
                    }
                    for (String s : subPaths) {
                        String finalPath = (API_PATH_PREFIX + classPath + (s.trim().startsWith("/") ? s.trim() : (s.trim().isEmpty() ? "" : "/" + s.trim()))).replaceAll("/+", "/");
                        ApiInfo info = new ApiInfo();
                        info.apiPath = (finalPath.isEmpty() ? "/" : finalPath);
                        info.methodName = method.getNameAsString(); info.isDeprecated = method.isAnnotationPresent("Deprecated") ? "Y" : "N";
                        info.controllerName = filePath.getFileName().toString(); info.repoPath = (REPO_NAME + "/" + relPath).replace("\\", "/");
                        info.git1 = git.get(0); info.git2 = git.get(1); info.git3 = git.get(2);
                        info.controllerComment = controllerComment; info.controllerRequestPropertyValue = controllerRequestProperty;
                        if (method.getComment().isPresent()) {
                            String full = method.getComment().get().getContent();
                            info.fullComment = full.replaceAll("\\r|\\n|\\*", " ").trim();
                            Matcher dM = Pattern.compile("@?(description|deprecation)[\\s:]*([^@\\n\\r*]+)", Pattern.CASE_INSENSITIVE).matcher(full);
                            info.descriptionTag = dM.find() ? dM.group(2).trim() : "-";
                        } else { info.fullComment = "-"; info.descriptionTag = "-"; }
                        info.requestPropertyValue = extractRequestPropertyFromNode(method);
                        info.apiOperationValue = extractAnnotationValue(method, "ApiOperation", "value");
                        apis.add(info);
                        log.append("\n      └ [Found] ").append(info.apiPath);
                    }
                }
            }
            // [v13.14 상세 로깅 추가] 매핑 어노테이션이 없어 누락되는 사유 명시 [cite: 2026-03-20]
            if (!hasMapping) {
                log.append("\n      └ [Skip] 매핑 어노테이션(RequestMapping 등) 미존재");
            }
        }
        return apis;
    }

    private static String evaluateExpression(Expression expr) {
        if (expr instanceof StringLiteralExpr) {
            return ((StringLiteralExpr) expr).getValue();
        } else if (expr instanceof BinaryExpr) {
            BinaryExpr be = (BinaryExpr) expr;
            if (be.getOperator() == BinaryExpr.Operator.PLUS) {
                return evaluateExpression(be.getLeft()) + evaluateExpression(be.getRight());
            }
        } else if (expr instanceof FieldAccessExpr || expr instanceof NameExpr) {
            String constName = expr.toString();
            return PATH_CONSTANTS_MAP.getOrDefault(constName, "{" + constName + "}");
        }
        return "";
    }

    private static List<String> getPathsFromAnn(AnnotationExpr ann) {
        List<String> paths = new ArrayList<>();
        Expression value = null;
        if (ann instanceof SingleMemberAnnotationExpr) {
            value = ((SingleMemberAnnotationExpr) ann).getMemberValue();
        } else if (ann instanceof NormalAnnotationExpr) {
            value = ((NormalAnnotationExpr) ann).getPairs().stream()
                    .filter(p -> p.getNameAsString().equals("value") || p.getNameAsString().equals("path"))
                    .map(MemberValuePair::getValue).findFirst().orElse(null);
        }

        if (value instanceof ArrayInitializerExpr) {
            for (Expression expr : ((ArrayInitializerExpr) value).getValues()) {
                String eval = evaluateExpression(expr);
                if (!eval.isEmpty()) paths.add(eval);
            }
        } else if (value != null) {
            String eval = evaluateExpression(value);
            if (!eval.isEmpty()) paths.add(eval);
        }
        return paths;
    }

    private static String extractRequestPropertyFromNode(com.github.javaparser.ast.nodeTypes.NodeWithAnnotations<?> node) {
        String title = extractValueFromNode(node, "RequestProperty", "title");
        return !"-".equals(title) ? title : extractValueFromNode(node, "RequestProperty", "value");
    }

    private static String extractValueFromNode(com.github.javaparser.ast.nodeTypes.NodeWithAnnotations<?> node, String annName, String attrName) {
        Optional<AnnotationExpr> ann = node.getAnnotationByName(annName);
        if (ann.isPresent() && ann.get() instanceof NormalAnnotationExpr) {
            return ((NormalAnnotationExpr) ann.get()).getPairs().stream().filter(p -> p.getNameAsString().equals(attrName)).map(p -> p.getValue().toString().replaceAll("\"", "")).findFirst().orElse("-");
        } else if (ann.isPresent() && ann.get() instanceof SingleMemberAnnotationExpr && "value".equals(attrName)) {
            return ((SingleMemberAnnotationExpr) ann.get()).getMemberValue().toString().replaceAll("\"", "");
        }
        return "-";
    }

    private static String extractAnnotationValue(MethodDeclaration method, String annName, String attrName) { return extractValueFromNode(method, annName, attrName); }

    private static List<ApiInfo> extractWithRegex(Path filePath, String relPath, List<String[]> git, StringBuilder log) {
        List<ApiInfo> apis = new ArrayList<>();
        try {
            String raw = new String(Files.readAllBytes(filePath), StandardCharsets.UTF_8);
            String clean = raw.replaceAll("(?s)/\\*.*?\\*/", " ").replaceAll("//.*", " ");
            Matcher cM_Main = Pattern.compile("/\\*\\*(.*?)\\*/", Pattern.DOTALL).matcher(raw);
            String controllerComment = cM_Main.find() ? cM_Main.group(1).replaceAll("\\r|\\n|\\*", " ").trim() : "-";
            String classPath = ""; Matcher cm = Pattern.compile("@RequestMapping\\s*\\(\\s*(?:(?:value|path)\\s*=\\s*)?\"([^\"]+)\"").matcher(clean);
            if (cm.find()) classPath = cm.group(1).trim();
            Matcher mMatcher = Pattern.compile("@(GetMapping|PostMapping|RequestMapping|PutMapping|DeleteMapping|PatchMapping)\\s*\\((.*?)\\)", Pattern.DOTALL).matcher(raw);
            while (mMatcher.find()) {
                String mappingType = mMatcher.group(1);
                String params = mMatcher.group(2);

                for (Map.Entry<String, String> entry : PATH_CONSTANTS_MAP.entrySet()) {
                    params = params.replace(entry.getKey(), "\"" + entry.getValue() + "\"");
                }
                params = params.replaceAll("\"\\s*\\+\\s*\"", "");

                Matcher mName = Pattern.compile("(?:public|private|protected)\\s+[\\w<>,\\s]+\\s+(\\w+)\\s*\\(").matcher(clean.substring(mMatcher.end(), Math.min(mMatcher.end() + 1000, clean.length())));
                if (mName.find()) {
                    String methodNameStr = mName.group(1);
                    // [v13.14 상세 로깅 추가] Regex 모드 진입 기록 [cite: 2026-03-20]
                    log.append("\n    * [Analyze-Regex] 메소드명: ").append(methodNameStr).append(" (").append(mappingType).append(")");

                    Matcher p = Pattern.compile("\"([^\"]+)\"").matcher(params);
                    boolean foundValidPath = false;
                    while (p.find()) {
                        String s = p.group(1).trim();
                        if (!s.contains("RequestMethod")) {
                            foundValidPath = true;
                            String finalPath = (API_PATH_PREFIX + classPath + (s.startsWith("/") ? s : (s.isEmpty() ? "" : "/" + s))).replaceAll("/+", "/");
                            ApiInfo info = new ApiInfo(); info.apiPath = (finalPath.isEmpty() ? "/" : finalPath);
                            info.methodName = methodNameStr; info.isDeprecated = clean.substring(Math.max(0, mMatcher.start() - 300), mMatcher.start()).contains("@Deprecated") ? "Y" : "N";
                            info.controllerName = filePath.getFileName().toString(); info.repoPath = (REPO_NAME + "/" + relPath).replace("\\", "/");
                            info.git1 = git.get(0); info.git2 = git.get(1); info.git3 = git.get(2);
                            info.controllerComment = controllerComment; String headArea = raw.substring(Math.max(0, mMatcher.start() - 1000), mMatcher.start());
                            Matcher cM = Pattern.compile("/\\*\\*(.*?)\\*/", Pattern.DOTALL).matcher(headArea);
                            if (cM.find()) { info.fullComment = cM.group(1).replaceAll("\\r|\\n|\\*", " ").trim();
                                Matcher dM = Pattern.compile("@?(description|deprecation)[\\s:]*([^@\\n\\r*]+)", Pattern.CASE_INSENSITIVE).matcher(cM.group(1));
                                info.descriptionTag = dM.find() ? dM.group(2).trim() : "-";
                            } else { info.fullComment = "-"; info.descriptionTag = "-"; }
                            apis.add(info);
                            log.append("\n      └ [Found-Regex] ").append(info.apiPath);
                        } else {
                            // [v13.14 상세 로깅 추가] 문자열 스킵 사유 [cite: 2026-03-20]
                            log.append("\n      └ [Skip-Regex] RequestMethod 포함 구문 스킵: ").append(s);
                        }
                    }
                    if (!foundValidPath) {
                        log.append("\n      └ [Skip-Regex] 유효한 문자열 경로 추출 실패");
                    }
                } else {
                    // [v13.14 상세 로깅 추가] 메소드 시그니처 매칭 실패 사유 [cite: 2026-03-20]
                    log.append("\n    * [Skip-Regex] 어노테이션(").append(mappingType).append(") 존재하나 메소드 매칭 실패");
                }
            }
        } catch (Exception ignored) {}
        return apis;
    }

    private static void addLog(String msg) { System.out.println(msg); if (logPath != null && !logPath.isEmpty()) { try (FileWriter fw = new FileWriter(logPath, true); PrintWriter pw = new PrintWriter(fw)) { pw.println(msg); } catch (IOException ignored) {} } }
    private static void saveInitialLogsToPath() { try (FileWriter fw = new FileWriter(logPath, false); PrintWriter pw = new PrintWriter(fw)) { pw.println("==============================================================="); pw.println("[START] " + REPO_NAME + " API 추출 및 Whatap 통합 시작 (v13.14)"); pw.println("==============================================================="); synchronized (RUNTIME_LOGS) { for (String l : RUNTIME_LOGS) pw.println(l); } } catch (IOException ignored) {} }
    private static void addExceptionLog(String title, Exception e) { StringWriter sw = new StringWriter(); e.printStackTrace(new PrintWriter(sw)); addLog("\n[ERROR] " + title + "\n" + sw.toString()); }

    private static List<String[]> getRecentGitHistories(String rel, String root, int c) {
        List<String[]> h = new ArrayList<>(); for (int i = 0; i < c; i++) h.add(new String[]{"-", "-", "No History"});
        try { Process p = new ProcessBuilder(GIT_BIN_PATH, "log", "-" + c, "--pretty=format:%as|%an|%s", "--", rel).directory(new File(root)).start();
            try (BufferedReader r = new BufferedReader(new InputStreamReader(p.getInputStream(), StandardCharsets.UTF_8))) {
                List<String> lines = new ArrayList<>(); String l; while ((l = r.readLine()) != null) lines.add(l); Collections.reverse(lines);
                for (int i = 0; i < lines.size(); i++) { String[] parts = lines.get(i).split("\\|", 3); h.set(i, new String[]{parts[0], parts[1], parts.length > 2 ? parts[2] : ""}); }
            }
            p.waitFor();
        } catch (Exception ignored) {}
        return h;
    }

    private static CellStyle createStyle(Workbook wb, Short bg, boolean bold, boolean center) {
        CellStyle s = wb.createCellStyle(); if (bg != null) { s.setFillForegroundColor(bg); s.setFillPattern(FillPatternType.SOLID_FOREGROUND); }
        s.setAlignment(center ? HorizontalAlignment.CENTER : HorizontalAlignment.LEFT); s.setVerticalAlignment(VerticalAlignment.CENTER);
        s.setBorderBottom(BorderStyle.THIN); s.setBorderTop(BorderStyle.THIN); s.setBorderLeft(BorderStyle.THIN); s.setBorderRight(BorderStyle.THIN);
        Font f = wb.createFont(); f.setBold(bold); s.setFont(f); return s;
    }

    static class ApiInfo {
        String apiPath, methodName, isDeprecated, controllerName, repoPath;
        String controllerComment, fullComment, descriptionTag, apiOperationValue, requestPropertyValue, controllerRequestPropertyValue;
        String[] git1, git2, git3; String getApiPath() { return apiPath; }
    }
}