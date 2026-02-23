import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
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
 * 프로젝트명: WhatapApiCounter (연간 통계 전수 추출 도구)
 * Version: 5.5 (보안 강화 버전)
 * [수정 사항]
 * 1. [보안 강화] WHATAP_URL, OUTPUT_DIR 기본값을 빈 값("")으로 수정하여 정보 노출 방지
 * 2. [로직 보존] v5.3의 모든 상세 주석, .log 파일 실시간 생성 및 parallelStream 수집 로직 유지
 * 3. [연동] generateExcelReport를 public으로 유지하여 ApiExcelExporter v11.2와 연동 보장
 */
public class WhatapApiCounter {

    // ==========================================================================================
    // [ 1. 시스템 설정 및 API 통신 변수 ]
    // ==========================================================================================

    /** 와탭 API 엔드포인트: 보안을 위해 기본값을 비웠습니다. config.properties에서 [WHATAP_URL]로 설정하세요. */
    private static String WHATAP_URL = "";

    /** Jackson Object Mapper: JSON 페이로드 생성 및 응답 데이터 파싱을 담당합니다. */
    private static final ObjectMapper MAPPER = new ObjectMapper().enable(SerializationFeature.INDENT_OUTPUT);

    /** 추출 활성화 플래그: config.properties의 'WHATAP_ENABLED' 값에 따릅니다. */
    private static boolean WHATAP_ENABLED = false;

    /** 통계 조회 시작일: 'YYYYMMDD' 형식의 문자열입니다. */
    private static String START_DATE = "";

    /** 통계 조회 종료일: 'YYYYMMDD' 형식의 문자열입니다. */
    private static String END_DATE = "";

    /** 와탭 세션 쿠키: 서버 보안 인증을 위한 키 값입니다. */
    private static String WHATAP_COOKIE = "";

    /** 결과물 저장 경로: 보안을 위해 기본값을 비웠습니다. config.properties에서 [OUTPUT_DIR]로 설정하세요. */
    private static String OUTPUT_DIR = "";

    /** 서비스 경로 필터 리스트: 특정 패턴(예: /app)을 가진 트랜잭션만 필터링합니다. */
    private static List<String> WHATAP_FILTERS = new ArrayList<>();

    /** 에이전트 그룹 ID들: 와탭에서 그룹핑된 대상 서버들의 식별 번호 목록입니다. */
    private static String WHATAP_OKINDS = "";

    /** 에이전트 그룹 명칭: 엑셀 파일명 및 로그 출력 시 식별 이름입니다. */
    private static String WHATAP_OKINDS_NAME = "";

    // ==========================================================================================
    // [ 2. 로그 파일 관리 및 기록 변수 ]
    // ==========================================================================================

    /** 로그 출력 스트림: 콘솔 내용을 실시간으로 .log 파일에 기록하는 객체입니다. */
    private static PrintWriter logWriter;

    /** 현재 실행 중인 세션의 로그 파일 절대 경로를 보관합니다. */
    private static String currentLogPath;

    // ==========================================================================================
    // [ 3. 데이터 관리 및 병렬 처리 변수 ]
    // ==========================================================================================

    /** API 통계 데이터 저장소: Key(경로) - Value(구간별 건수 배열) 구조의 Thread-safe Map입니다. */
    private static final Map<String, long[]> STATS_MAP = new ConcurrentHashMap<>();

    /** 수집 구간 리스트: 수집 기간을 10일 단위로 쪼갠 세부 정보들의 모음입니다. */
    private static final List<FetchSegment> SEGMENTS = new ArrayList<>();

    public static class FetchSegment {
        public String label;
        public long stime;
        public long etime;
        public String monthKey;
    }

    // ==========================================================================================

    /** [연동 인터페이스] 외부 클래스에서 호출 시 수집된 통계 Map 데이터를 반환합니다. */
    public static Map<String, long[]> getApiStats() {
        STATS_MAP.clear();
        SEGMENTS.clear();

        if (START_DATE.isEmpty()) loadConfig();

        if (!WHATAP_ENABLED) return STATS_MAP;

        generateSegments();
        fetchBatchData();

        return STATS_MAP;
    }

    public static void main(String[] args) {
        loadConfig();

        LocalDateTime execStartTime = LocalDateTime.now();
        String timestamp = execStartTime.format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        DateTimeFormatter logFmt = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

        initLogFile(timestamp);

        addLog("===============================================================");
        addLog("[START] WhatapApiCounter v5.5 실행 시작: " + execStartTime.format(logFmt));
        addLog("===============================================================");

        getApiStats();

        if (WHATAP_ENABLED && !STATS_MAP.isEmpty()) {
            generateExcelReport(timestamp);
        }

        LocalDateTime execEndTime = LocalDateTime.now();
        addLog("\n===============================================================");
        addLog("[FINISH] WhatapApiCounter 실행 종료: " + execEndTime.format(logFmt));
        addLog("[RESULT] 총 소요 시간: " + Duration.between(execStartTime, execEndTime).getSeconds() + "초");
        addLog("[RESULT] 총 수집 고유 API: " + STATS_MAP.size() + "건");
        addLog("[LOG_FILE] 로그 확인 경로: " + currentLogPath);
        addLog("===============================================================");

        if (logWriter != null) logWriter.close();
    }

    private static synchronized void addLog(String msg) {
        System.out.println(msg);
        if (logWriter != null) {
            logWriter.println(msg);
            logWriter.flush();
        }
    }

    private static void initLogFile(String ts) {
        try {
            if (OUTPUT_DIR.isEmpty()) return;
            File dir = new File(OUTPUT_DIR);
            if (!dir.exists()) dir.mkdirs();

            String fileName = "Whatap_통계_추출로그_" + ts + ".log";
            File logFile = new File(dir, fileName);
            currentLogPath = logFile.getAbsolutePath();
            logWriter = new PrintWriter(new BufferedWriter(new OutputStreamWriter(new FileOutputStream(logFile), StandardCharsets.UTF_8)));
        } catch (IOException e) {
            System.err.println("로그 파일 생성 중 오류: " + e.getMessage());
        }
    }

    /** [지적 반영] 보안 변수들을 프로퍼티에서 로드하는 로직으로 강화 [cite: 2026-02-23] */
    private static void loadConfig() {
        Properties prop = new Properties();
        File configFile = new File("config.properties");
        if (configFile.exists()) {
            try (InputStream is = new FileInputStream(configFile)) {
                prop.load(is);
                WHATAP_ENABLED = "Y".equalsIgnoreCase(prop.getProperty("WHATAP_ENABLED", "N"));
                START_DATE = prop.getProperty("START_DATE", "").trim();
                END_DATE = prop.getProperty("END_DATE", "").trim();
                WHATAP_COOKIE = prop.getProperty("WHATAP_COOKIE", "").trim();
                if (WHATAP_COOKIE.startsWith("\"")) WHATAP_COOKIE = WHATAP_COOKIE.substring(1, WHATAP_COOKIE.length()-1);
                WHATAP_OKINDS = prop.getProperty("WHATAP_OKINDS", "").trim();
                WHATAP_OKINDS_NAME = prop.getProperty("WHATAP_OKINDS_NAME", "Unknown").trim();

                // [보안 지적 반영] URL과 저장경로를 프로퍼티에서 로드
                WHATAP_URL = prop.getProperty("WHATAP_URL", "").trim();
                OUTPUT_DIR = prop.getProperty("OUTPUT_DIR", "").trim();

                String fProp = prop.getProperty("WHATAP_FILTER", "").trim();
                WHATAP_FILTERS = Arrays.stream(fProp.split(",")).map(String::trim).filter(s -> !s.isEmpty()).collect(Collectors.toList());
                if (WHATAP_FILTERS.isEmpty()) WHATAP_FILTERS.add("");

                addLog("[LOG] 설정값 로드 상세 내역:");
                addLog("  > WHATAP_URL     : " + (WHATAP_URL.isEmpty() ? "MISSING!" : WHATAP_URL));
                addLog("  > OKINDS_NAME    : " + WHATAP_OKINDS_NAME);
                addLog("  > OKINDS_ID      : " + WHATAP_OKINDS);
                addLog("  > START_DATE     : " + START_DATE);
                addLog("  > END_DATE       : " + END_DATE);
                addLog("  > WHATAP_FILTER  : " + WHATAP_FILTERS);
                addLog("  > OUTPUT_DIR     : " + (OUTPUT_DIR.isEmpty() ? "MISSING!" : OUTPUT_DIR));
                addLog("  > WHATAP_ENABLED : " + WHATAP_ENABLED);
                addLog("---------------------------------------------------------------");
            } catch (IOException e) { e.printStackTrace(); }
        }
    }

    private static void generateSegments() {
        LocalDate startLimit = LocalDate.parse(START_DATE, DateTimeFormatter.ofPattern("yyyyMMdd"));
        LocalDate endLimit = LocalDate.parse(END_DATE, DateTimeFormatter.ofPattern("yyyyMMdd"));
        LocalDate cur = startLimit.withDayOfMonth(1);
        while (!cur.isAfter(endLimit)) {
            String mKey = cur.format(DateTimeFormatter.ofPattern("yy.MM"));
            addValidSegment(cur.withDayOfMonth(1), cur.withDayOfMonth(10), startLimit, endLimit, mKey);
            addValidSegment(cur.withDayOfMonth(11), cur.withDayOfMonth(20), startLimit, endLimit, mKey);
            addValidSegment(cur.withDayOfMonth(21), cur.withDayOfMonth(cur.lengthOfMonth()), startLimit, endLimit, mKey);
            cur = cur.plusMonths(1);
        }
    }

    private static void addValidSegment(LocalDate s, LocalDate e, LocalDate limitS, LocalDate limitE, String monthKey) {
        LocalDate actualS = s.isBefore(limitS) ? limitS : s;
        LocalDate actualE = e.isAfter(limitE) ? limitE : e;
        if (!actualS.isAfter(actualE)) {
            FetchSegment seg = new FetchSegment();
            seg.label = actualS.format(DateTimeFormatter.ofPattern("yyyy-MM-dd")) + "~" + actualE.getDayOfMonth();
            seg.stime = actualS.atStartOfDay(ZoneId.systemDefault()).toInstant().toEpochMilli();
            seg.etime = actualE.plusDays(1).atStartOfDay(ZoneId.systemDefault()).toInstant().toEpochMilli();
            seg.monthKey = monthKey;
            SEGMENTS.add(seg);
        }
    }

    private static void fetchBatchData() {
        ExecutorService executor = Executors.newFixedThreadPool(3);
        List<CompletableFuture<Void>> futures = new ArrayList<>();
        for (int i = 0; i < SEGMENTS.size(); i++) {
            final int segIdx = i;
            final FetchSegment seg = SEGMENTS.get(i);
            for (String filter : WHATAP_FILTERS) {
                futures.add(CompletableFuture.runAsync(() -> requestWithDetailedFetch(seg.stime, seg.etime, segIdx, filter), executor));
            }
        }
        CompletableFuture.allOf(futures.toArray(new CompletableFuture[0])).join();
        executor.shutdown();
    }

    private static void requestWithDetailedFetch(long stime, long etime, int segIdx, String filter) {
        String label = SEGMENTS.get(segIdx).label;
        try {
            if (WHATAP_URL.isEmpty()) { addLog("  - [SKIP] URL이 설정되지 않았습니다."); return; }

            String jsonPayload = String.format(
                    "{\n" +
                            "  \"type\": \"stat\",\n" +
                            "  \"path\": \"ap\",\n" +
                            "  \"pcode\": 8,\n" +
                            "  \"params\": {\n" +
                            "    \"stime\": %d,\n" +
                            "    \"etime\": %d,\n" +
                            "    \"ptotal\": 100,\n" +
                            "    \"skip\": 0,\n" +
                            "    \"psize\": 10000,\n" +
                            "    \"filter\": { \"service\": \"%s\" },\n" +
                            "    \"okinds\": [%s],\n" +
                            "    \"order\": \"countTotal\",\n" +
                            "    \"type\": \"service\"\n" +
                            "  },\n" +
                            "  \"stime\": %d,\n" +
                            "  \"etime\": %d\n" +
                            "}", stime, etime, filter, WHATAP_OKINDS, stime, etime
            );

            addLog("\n>>> [HTTP REQUEST] 구간: " + label + " (필터: " + filter + ")");
            addLog("  Payload: " + jsonPayload);

            HttpClient client = HttpClient.newBuilder().connectTimeout(Duration.ofSeconds(20)).build();
            HttpRequest request = HttpRequest.newBuilder().uri(URI.create(WHATAP_URL)).header("Content-Type", "application/json").header("Cookie", WHATAP_COOKIE).POST(HttpRequest.BodyPublishers.ofString(jsonPayload, StandardCharsets.UTF_8)).build();

            HttpResponse<String> response = client.send(request, HttpResponse.BodyHandlers.ofString());
            String body = response.body();

            if (response.statusCode() == 200) {
                JsonNode root = MAPPER.readTree(body).path("records");
                int count = 0;
                if (root.isArray()) {
                    for (JsonNode n : root) {
                        String svc = n.path("service").asText();
                        long cnt = n.path("count").asLong();
                        long[] stats = STATS_MAP.computeIfAbsent(svc, k -> new long[SEGMENTS.size() + 10]);
                        synchronized (stats) { stats[segIdx] += cnt; }
                        count++;
                    }
                }
                addLog("  - [INFO] " + label + " 수집 완료 (" + count + "건)");
            }
        } catch (Exception e) {
            addLog("  - [ERROR] " + label + " 통계 수집 중 예외: " + e.getMessage());
        }
    }

    public static void generateExcelReport(String ts) {
        if (OUTPUT_DIR.isEmpty()) { System.err.println("[ERROR] OUTPUT_DIR이 비어있어 엑셀을 생성할 수 없습니다."); return; }

        String fileName = String.format("Whatap 통계 추출결과_(v5.3)_(%s)_(%s~%s)_(%s).xlsx", WHATAP_OKINDS_NAME, START_DATE, END_DATE, ts);
        File file = new File(OUTPUT_DIR, fileName);
        if (!file.getParentFile().exists()) file.getParentFile().mkdirs();

        try (SXSSFWorkbook wb = new SXSSFWorkbook(100); FileOutputStream fos = new FileOutputStream(file)) {
            Sheet s = wb.createSheet("Whatap_" + WHATAP_OKINDS_NAME);
            s.createFreezePane(2, 1);

            DataFormat df = wb.createDataFormat();
            short numFmt = df.getFormat("#,##0");
            Font hFont = wb.createFont(); hFont.setBold(true);

            CellStyle grayT = createHeaderStyle(wb, hFont, IndexedColors.GREY_25_PERCENT.getIndex());
            grayT.setBorderRight(BorderStyle.THICK);
            CellStyle harmonyLightStyle = createHeaderStyle(wb, hFont, IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
            CellStyle harmonyDarkStyle = createHeaderStyle(wb, hFont, IndexedColors.CORNFLOWER_BLUE.getIndex());
            harmonyDarkStyle.setBorderRight(BorderStyle.THICK);

            CellStyle apiS = wb.createCellStyle(); apiS.setBorderRight(BorderStyle.THICK);
            CellStyle numS = wb.createCellStyle(); numS.setDataFormat(numFmt);
            CellStyle numThickS = wb.createCellStyle(); numThickS.setDataFormat(numFmt); numThickS.setBorderRight(BorderStyle.THICK);

            Row hr = s.createRow(0);
            hr.createCell(0).setCellValue("API(트랜잭션)"); hr.getCell(0).setCellStyle(grayT);
            hr.createCell(1).setCellValue("전체 총합계"); hr.getCell(1).setCellStyle(grayT);

            int colIdx = 2; String lastMonth = "";
            Map<String, List<Integer>> monthCols = new LinkedHashMap<>();
            for (int i = 0; i < SEGMENTS.size(); i++) {
                FetchSegment seg = SEGMENTS.get(i);
                if (!seg.monthKey.equals(lastMonth) && !lastMonth.isEmpty()) {
                    Cell c = hr.createCell(colIdx++); c.setCellValue(lastMonth + " 월 합계"); c.setCellStyle(harmonyDarkStyle);
                }
                Cell c = hr.createCell(colIdx); c.setCellValue(seg.label); c.setCellStyle(harmonyLightStyle);
                monthCols.computeIfAbsent(seg.monthKey, k -> new ArrayList<>()).add(colIdx++);
                lastMonth = seg.monthKey;
            }
            Cell cLast = hr.createCell(colIdx++); cLast.setCellValue(lastMonth + " 월 합계"); cLast.setCellStyle(harmonyDarkStyle);

            List<Map.Entry<String, long[]>> sorted = STATS_MAP.entrySet().stream()
                    .sorted((e1, e2) -> Long.compare(calculateGrandTotal(e2.getValue()), calculateGrandTotal(e1.getValue())))
                    .collect(Collectors.toList());

            int rIdx = 1;
            for (Map.Entry<String, long[]> entry : sorted) {
                Row r = s.createRow(rIdx++);
                r.createCell(0).setCellValue(entry.getKey()); r.getCell(0).setCellStyle(apiS);
                r.createCell(1).setCellValue(calculateGrandTotal(entry.getValue())); r.getCell(1).setCellStyle(numThickS);

                int dCol = 2; int ptr = 0;
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
            s.setColumnWidth(0, 18000); s.setColumnWidth(1, 6500);
            for (int i = 2; i < colIdx; i++) s.setColumnWidth(i, 5500);

            wb.write(fos);
            wb.dispose();

            addLog("\n[SUCCESS] 통계 엑셀 생성 완료");
            addLog("  > 저장 위치 : " + file.getParent());
            addLog("  > 파 일 명  : " + file.getName());
            addLog("  > 전체 경로 : " + file.getAbsolutePath());

        } catch (Exception e) { e.printStackTrace(); }
    }

    private static CellStyle createHeaderStyle(Workbook wb, Font f, short color) {
        CellStyle st = wb.createCellStyle();
        st.setFillForegroundColor(color); st.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        st.setAlignment(HorizontalAlignment.CENTER); st.setBorderBottom(BorderStyle.THIN); st.setFont(f);
        return st;
    }

    private static long calculateGrandTotal(long[] stats) {
        long sum = 0; if(stats != null) for (long v : stats) sum += v; return sum;
    }
}