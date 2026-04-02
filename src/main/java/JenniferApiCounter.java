import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.*;
import java.net.URI;
import java.net.http.*;
import java.nio.charset.StandardCharsets;
import java.time.*;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.*;
import java.util.stream.Collectors;

/**
 * 프로젝트명: GenericApmCounter (범용 APM 통계 추출 도구)
 * Version: 1.0
 *
 * WhatapApiCounter와 동일한 Excel/Log 출력을 생성하되,
 * Whatap이 아닌 다른 APM의 Open API와 연동합니다.
 *
 * ──────────────────────────────────────────────────────────────
 * [APM 추가 방법 - 2단계]
 *
 * STEP 1. config.properties 설정:
 *   JENNIFER_ENABLED=Y
 *   JENNIFER_URL=https://your-apm-host/api/v1/stats    ← 대상 APM 엔드포인트
 *   JENNIFER_KEY=your-api-key-here                      ← 인증 키
 *   JENNIFER_AUTH_SCHEME=Bearer                         ← 인증 방식 (Bearer / Basic / ApiKey)
 *   JENNIFER_DISPLAY_NAME=MyService                     ← 엑셀 파일명에 표시할 이름
 *   JENNIFER_FILTER=/api,/app                           ← 서비스 경로 필터 (없으면 전체)
 *   START_DATE=20250101                            ← WhatapApiCounter와 공유 가능
 *   END_DATE=20251231
 *   OUTPUT_DIR=/path/to/output
 *
 * STEP 2. buildRequestBody() / parseResponse() 를 대상 APM 스펙에 맞게 수정:
 *   - buildRequestBody(): APM에 보낼 요청 Body(JSON/Query 등) 생성
 *   - parseResponse():    APM 응답에서 Map<서비스경로, 호출건수> 추출
 * ──────────────────────────────────────────────────────────────
 */
public class JenniferApiCounter {

    // ── 공통 설정 (WhatapApiCounter와 동일한 키 공유 가능) ──────────────────
    private static String START_DATE  = "";
    private static String END_DATE    = "";
    private static String OUTPUT_DIR  = "";

    // ── APM별 설정 ──────────────────────────────────────────────────────────
    private static boolean JENNIFER_ENABLED      = false;
    private static String  JENNIFER_URL          = "";
    /** API 인증 키 (Bearer 토큰, API Key 등) */
    private static String  JENNIFER_KEY          = "";
    /** Authorization 헤더 scheme. Bearer / Basic / ApiKey 중 선택. 기본값: Bearer */
    private static String  JENNIFER_AUTH_SCHEME  = "Bearer";
    /** 엑셀 파일명/시트명에 표시할 서비스 이름 */
    private static String  JENNIFER_DISPLAY_NAME = "Unknown";
    /** 도메인 ID */
    private static String  JENNIFER_DOMAIN_ID = "";
    /** 서비스 경로 필터 목록 (없으면 [""] 로 전체 수집) */
    private static List<String> JENNIFER_FILTERS = new ArrayList<>();

    // ── 내부 상태 ────────────────────────────────────────────────────────────
    private static final Map<String, long[]> STATS_MAP = new ConcurrentHashMap<>();
    private static final List<FetchSegment>  SEGMENTS  = new ArrayList<>();
    private static final ObjectMapper MAPPER = new ObjectMapper();

    private static PrintWriter logWriter;
    private static String      currentLogPath;

    public static class FetchSegment {
        public String label;
        public long   stime;
        public long   etime;
        public String monthKey;
    }

    // ════════════════════════════════════════════════════════════════════════
    // PUBLIC API (WhatapApiCounter.getApiStats()와 동일한 시그니처)
    // ════════════════════════════════════════════════════════════════════════

    public static Map<String, long[]> getApiStats() {
        STATS_MAP.clear();
        SEGMENTS.clear();

        if (START_DATE.isEmpty()) loadConfig();
        if (!JENNIFER_ENABLED) return STATS_MAP;

        generateSegments();
        fetchBatchData();

        return STATS_MAP;
    }

    public static void main(String[] args) {
        loadConfig();

        LocalDateTime execStartTime = LocalDateTime.now();
        String timestamp = execStartTime.format(DateTimeFormatter.ofPattern("yyyy-MM-dd'_추출'"));
        DateTimeFormatter logFmt = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

        initLogFile(timestamp);

        addLog("===============================================================");
        addLog("[START] GenericApmCounter v1.0 실행 시작: " + execStartTime.format(logFmt));
        addLog("===============================================================");

        getApiStats();

        if (JENNIFER_ENABLED && !STATS_MAP.isEmpty()) {
            generateExcelReport(timestamp);
        }

        LocalDateTime execEndTime = LocalDateTime.now();
        addLog("\n===============================================================");
        addLog("[FINISH] GenericApmCounter 실행 종료: " + execEndTime.format(logFmt));
        addLog("[RESULT] 총 소요 시간: " + Duration.between(execStartTime, execEndTime).getSeconds() + "초");
        addLog("[RESULT] 총 수집 고유 API: " + STATS_MAP.size() + "건");
        addLog("[LOG_FILE] 로그 확인 경로: " + currentLogPath);
        addLog("===============================================================");

        if (logWriter != null) logWriter.close();
    }

    // ════════════════════════════════════════════════════════════════════════
    // STEP 2-A ★ 수정 포인트: Query String 생성
    // 대상 APM의 API 스펙에 맞게 이 메서드를 수정하세요.
    // ════════════════════════════════════════════════════════════════════════

    /**
     * APM GET 요청에 붙일 Query String을 생성합니다. (? 는 포함하지 않음)
     *
     * @param stime  조회 시작 (epoch ms)
     * @param etime  조회 종료 (epoch ms)
     * @return "key=value&key2=value2" 형태의 query string
     *
     * ── 예시: Datadog APM ─────────────────────────────────────────────
     * "from=" + stime + "&to=" + etime + "&filter[service]=" + encode(filter)
     *
     * ── 예시: Dynatrace ───────────────────────────────────────────────
     * "from=" + stime + "&to=" + etime + "&entitySelector=type(SERVICE)"
     * ──────────────────────────────────────────────────────────────────
     */
    private static String buildQueryString(long stime, long etime) {
        // ↓↓↓ 여기를 대상 APM 스펙에 맞게 수정 ↓↓↓
        return "token=" + JENNIFER_KEY + "&domain_id=" + JENNIFER_DOMAIN_ID + "&startTime=" + stime + "&endTime=" + etime;
        // ↑↑↑ 여기를 대상 APM 스펙에 맞게 수정 ↑↑↑
    }

    // ════════════════════════════════════════════════════════════════════════
    // STEP 2-B ★ 수정 포인트: 응답 파싱
    // 대상 APM의 응답 JSON 구조에 맞게 이 메서드를 수정하세요.
    // ════════════════════════════════════════════════════════════════════════

    /**
     * APM 응답 Body를 파싱하여 Map&lt;서비스경로, 호출건수&gt;를 반환합니다.
     *
     * @param responseBody HTTP 응답 Body 문자열
     * @return Map&lt;서비스 경로, 호출 건수&gt;
     *
     * ── 예시: Elastic APM 응답 ──────────────────────────────────────────
     * { "hits": [ { "key": "/api/v1/users", "doc_count": 500 }, ... ] }
     * → serviceField = "key", countField = "doc_count", rootPath = "hits"
     *
     * ── 예시: Dynatrace 응답 ────────────────────────────────────────────
     * { "result": [ { "metricId": "/api/users", "data": [{"values":[100]}] } ] }
     * → 커스텀 파싱 필요
     * ─────────────────────────────────────────────────────────────────────
     */
    private static Map<String, Long> parseResponse(String responseBody) throws Exception {
        Map<String, Long> result = new HashMap<>();

        // ↓↓↓ 여기를 대상 APM 응답 스펙에 맞게 수정 ↓↓↓
        JsonNode root = MAPPER.readTree(responseBody);

        // 예시: { "records": [ { "service": "/api/...", "count": 123 } ] }
        // → Whatap 응답과 동일한 구조라면 그대로 사용 가능
        String rootPath    = "result";  // ← 응답의 배열 필드명
        String serviceField = "name"; // ← 서비스 경로 필드명
        String countField   = "calls";   // ← 호출 건수 필드명

        JsonNode records = root.path(rootPath);
        if (records.isArray()) {
            for (JsonNode node : records) {
                String svc = node.path(serviceField).asText();
                long   cnt = node.path(countField).asLong();
                if (!svc.isBlank()) result.merge(svc, cnt, Long::sum);
            }
        }
        // ↑↑↑ 여기를 대상 APM 응답 스펙에 맞게 수정 ↑↑↑

        return result;
    }

    // ════════════════════════════════════════════════════════════════════════
    // 공통 처리 로직 (수정 불필요)
    // ════════════════════════════════════════════════════════════════════════

    private static void fetchBatchData() {
        ExecutorService executor = Executors.newFixedThreadPool(3);
        List<CompletableFuture<Void>> futures = new ArrayList<>();

        for (int i = 0; i < SEGMENTS.size(); i++) {
            final int segIdx = i;
            final FetchSegment seg = SEGMENTS.get(i);

            futures.add(CompletableFuture.runAsync(() -> requestWithDetailedFetch(seg, segIdx), executor));
        }

        CompletableFuture.allOf(futures.toArray(new CompletableFuture[0])).join();
        executor.shutdown();
    }

    private static void requestWithDetailedFetch(FetchSegment seg, int segIdx) {
        try {
            if (JENNIFER_URL.isEmpty()) {
                addLog("  - [SKIP] JENNIFER_URL이 설정되지 않았습니다.");
                return;
            }

            String queryString = buildQueryString(seg.stime, seg.etime);
            String requestUrl  = JENNIFER_URL + "?" + queryString;
            addLog("  URL: " + requestUrl);

            HttpClient client = HttpClient.newBuilder()
                    .connectTimeout(Duration.ofSeconds(20))
                    .build();

            HttpRequest request = HttpRequest.newBuilder()
                    .uri(URI.create(requestUrl))
                    .header("Content-Type", "application/json")
                    .header("Authorization", JENNIFER_AUTH_SCHEME + " " + JENNIFER_KEY)
                    .GET()

                    .build();

            HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());

            if (response.statusCode() == 200) {
                Map<String, Long> stats = parseResponse(response.body());
                for (Map.Entry<String, Long> entry : stats.entrySet()) {
                    long[] arr = STATS_MAP.computeIfAbsent(
                            entry.getKey(), k -> new long[SEGMENTS.size() + 10]);
                    synchronized (arr) { arr[segIdx] += entry.getValue(); }
                }
                addLog("  - [INFO] " + seg.label + " 수집 완료 (" + stats.size() + "건)");
            } else {
                addLog("  - [WARN] HTTP " + response.statusCode() + ": " + response.body());
            }

        } catch (Exception e) {
            addLog("  - [ERROR] " + seg.label + " 통계 수집 중 예외: " + e.getMessage());
        }
    }

    private static void generateSegments() {
        LocalDate startLimit = LocalDate.parse(START_DATE, DateTimeFormatter.ofPattern("yyyyMMdd"));
        LocalDate endLimit   = LocalDate.parse(END_DATE,   DateTimeFormatter.ofPattern("yyyyMMdd"));
        LocalDate cur = startLimit.withDayOfMonth(1);
        while (!cur.isAfter(endLimit)) {
            String mKey = cur.format(DateTimeFormatter.ofPattern("yy.MM"));
            addValidSegment(cur, cur.plusDays(1), startLimit, endLimit, mKey);
            cur = cur.plusDays(1);
        }
    }

    private static void addValidSegment(LocalDate s, LocalDate e,
                                        LocalDate limitS, LocalDate limitE, String monthKey) {
        LocalDate actualS = s.isBefore(limitS) ? limitS : s;
        LocalDate actualE = e.isAfter(limitE)  ? limitE : e;
        if (!actualS.isAfter(actualE)) {
            FetchSegment seg = new FetchSegment();
            seg.label    = actualS.format(DateTimeFormatter.ofPattern("yyyy-MM-dd"));
            seg.stime    = actualS.atStartOfDay(ZoneId.systemDefault()).toInstant().toEpochMilli();
            seg.etime    = actualE.atStartOfDay(ZoneId.systemDefault()).toInstant().toEpochMilli();
            seg.monthKey = monthKey;
            SEGMENTS.add(seg);
        }
    }

    public static void generateExcelReport(String ts) {
        if (OUTPUT_DIR.isEmpty()) {
            System.err.println("[ERROR] OUTPUT_DIR이 비어있어 엑셀을 생성할 수 없습니다.");
            return;
        }

        String fileName = String.format("APM통계_(%s)_(%s~%s)_(%s).xlsx",
                JENNIFER_DISPLAY_NAME, START_DATE, END_DATE, ts);
        File file = new File(OUTPUT_DIR, fileName);
        if (!file.getParentFile().exists()) file.getParentFile().mkdirs();

        try (SXSSFWorkbook wb = new SXSSFWorkbook(100);
             FileOutputStream fos = new FileOutputStream(file)) {

            Sheet s = wb.createSheet(JENNIFER_DISPLAY_NAME);
            s.createFreezePane(2, 1);

            DataFormat df   = wb.createDataFormat();
            short numFmt    = df.getFormat("#,##0");
            Font  hFont     = wb.createFont();
            hFont.setBold(true);

            CellStyle grayT          = createHeaderStyle(wb, hFont, IndexedColors.GREY_25_PERCENT.getIndex());
            grayT.setBorderRight(BorderStyle.THICK);
            CellStyle lightStyle     = createHeaderStyle(wb, hFont, IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
            CellStyle darkStyle      = createHeaderStyle(wb, hFont, IndexedColors.CORNFLOWER_BLUE.getIndex());
            darkStyle.setBorderRight(BorderStyle.THICK);

            CellStyle apiS     = wb.createCellStyle(); apiS.setBorderRight(BorderStyle.THICK);
            CellStyle numS     = wb.createCellStyle(); numS.setDataFormat(numFmt);
            CellStyle numThickS = wb.createCellStyle(); numThickS.setDataFormat(numFmt); numThickS.setBorderRight(BorderStyle.THICK);

            // 헤더 행
            Row hr = s.createRow(0);
            hr.createCell(0).setCellValue("API(트랜잭션)"); hr.getCell(0).setCellStyle(grayT);
            hr.createCell(1).setCellValue("전체 총합계");   hr.getCell(1).setCellStyle(grayT);

            int colIdx = 2;
            String lastMonth = "";
            Map<String, List<Integer>> monthCols = new LinkedHashMap<>();

            for (int i = 0; i < SEGMENTS.size(); i++) {
                FetchSegment seg = SEGMENTS.get(i);
                if (!seg.monthKey.equals(lastMonth) && !lastMonth.isEmpty()) {
                    Cell c = hr.createCell(colIdx++);
                    c.setCellValue(lastMonth + " 월 합계"); c.setCellStyle(darkStyle);
                }
                Cell c = hr.createCell(colIdx);
                c.setCellValue(seg.label); c.setCellStyle(lightStyle);
                monthCols.computeIfAbsent(seg.monthKey, k -> new ArrayList<>()).add(colIdx++);
                lastMonth = seg.monthKey;
            }
            Cell cLast = hr.createCell(colIdx++);
            cLast.setCellValue(lastMonth + " 월 합계"); cLast.setCellStyle(darkStyle);

            // 데이터 행 (총합 내림차순 정렬)
            List<Map.Entry<String, long[]>> sorted = STATS_MAP.entrySet().stream()
                    .sorted((e1, e2) -> Long.compare(calculateGrandTotal(e2.getValue()), calculateGrandTotal(e1.getValue())))
                    .collect(Collectors.toList());

            int rIdx = 1;
            for (Map.Entry<String, long[]> entry : sorted) {
                Row r = s.createRow(rIdx++);
                r.createCell(0).setCellValue(entry.getKey()); r.getCell(0).setCellStyle(apiS);
                r.createCell(1).setCellValue(calculateGrandTotal(entry.getValue())); r.getCell(1).setCellStyle(numThickS);

                int dCol = 2, ptr = 0;
                for (String m : monthCols.keySet()) {
                    long mSum = 0;
                    for (int ignored : monthCols.get(m)) {
                        long v = entry.getValue()[ptr++];
                        Cell c = r.createCell(dCol++); c.setCellValue(v); c.setCellStyle(numS);
                        mSum += v;
                    }
                    Cell cS = r.createCell(dCol++); cS.setCellValue(mSum); cS.setCellStyle(numThickS);
                }
            }

            s.setColumnWidth(0, 18000);
            s.setColumnWidth(1, 6500);
            for (int i = 2; i < colIdx; i++) s.setColumnWidth(i, 5500);

            wb.write(fos);
            wb.dispose();

            addLog("\n[SUCCESS] 통계 엑셀 생성 완료");
            addLog("  > 저장 위치 : " + file.getParent());
            addLog("  > 파 일 명  : " + file.getName());
            addLog("  > 전체 경로 : " + file.getAbsolutePath());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ════════════════════════════════════════════════════════════════════════
    // 설정 로드 / 로그
    // ════════════════════════════════════════════════════════════════════════

    private static void loadConfig() {
        Properties prop = new Properties();
        File configFile = new File("config.properties");
        if (!configFile.exists()) return;

        try (InputStream is = new FileInputStream(configFile)) {
            prop.load(is);

            // 공통 설정
            START_DATE = prop.getProperty("START_DATE", "").trim();
            END_DATE   = prop.getProperty("END_DATE",   "").trim();
            OUTPUT_DIR = prop.getProperty("OUTPUT_DIR", "").trim();

            // APM별 설정
            JENNIFER_ENABLED      = "Y".equalsIgnoreCase(prop.getProperty("JENNIFER_ENABLED", "N"));
            JENNIFER_URL          = prop.getProperty("JENNIFER_URL",          "").trim();
            JENNIFER_KEY          = prop.getProperty("JENNIFER_KEY",          "").trim();
            JENNIFER_AUTH_SCHEME  = prop.getProperty("JENNIFER_AUTH_SCHEME",  "Bearer").trim();
            JENNIFER_DISPLAY_NAME = prop.getProperty("JENNIFER_DISPLAY_NAME", "Unknown").trim();
            JENNIFER_DOMAIN_ID = prop.getProperty("JENNIFER_DOMAIN_ID", "Unknown").trim();

            String fProp = prop.getProperty("JENNIFER_FILTER", "").trim();
            JENNIFER_FILTERS = Arrays.stream(fProp.split(","))
                    .map(String::trim).filter(s -> !s.isEmpty())
                    .collect(Collectors.toList());
            if (JENNIFER_FILTERS.isEmpty()) JENNIFER_FILTERS.add("");

            addLog("[LOG] 설정값 로드 상세 내역:");
            addLog("  > JENNIFER_URL          : " + (JENNIFER_URL.isEmpty() ? "MISSING!" : JENNIFER_URL));
            addLog("  > JENNIFER_AUTH_SCHEME  : " + JENNIFER_AUTH_SCHEME);
            addLog("  > JENNIFER_DISPLAY_NAME : " + JENNIFER_DISPLAY_NAME);
            addLog("  > JENNIFER_FILTER       : " + JENNIFER_FILTERS);
            addLog("  > START_DATE       : " + START_DATE);
            addLog("  > END_DATE         : " + END_DATE);
            addLog("  > OUTPUT_DIR       : " + (OUTPUT_DIR.isEmpty() ? "MISSING!" : OUTPUT_DIR));
            addLog("  > JENNIFER_ENABLED      : " + JENNIFER_ENABLED);
            addLog("---------------------------------------------------------------");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void initLogFile(String ts) {
        try {
            if (OUTPUT_DIR.isEmpty()) return;
            File dir = new File(OUTPUT_DIR);
            if (!dir.exists()) dir.mkdirs();
            String fileName = "JENNIFER_통계_추출로그_" + ts + ".log";
            File logFile = new File(dir, fileName);
            currentLogPath = logFile.getAbsolutePath();
            logWriter = new PrintWriter(new BufferedWriter(
                    new OutputStreamWriter(new FileOutputStream(logFile), StandardCharsets.UTF_8)));
        } catch (IOException e) {
            System.err.println("로그 파일 생성 중 오류: " + e.getMessage());
        }
    }

    private static synchronized void addLog(String msg) {
        System.out.println(msg);
        if (logWriter != null) { logWriter.println(msg); logWriter.flush(); }
    }

    private static CellStyle createHeaderStyle(Workbook wb, Font f, short color) {
        CellStyle st = wb.createCellStyle();
        st.setFillForegroundColor(color);
        st.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        st.setAlignment(HorizontalAlignment.CENTER);
        st.setBorderBottom(BorderStyle.THIN);
        st.setFont(f);
        return st;
    }

    private static long calculateGrandTotal(long[] stats) {
        long sum = 0;
        if (stats != null) for (long v : stats) sum += v;
        return sum;
    }
}