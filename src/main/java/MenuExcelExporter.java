import com.fasterxml.jackson.core.JsonParser;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * 프로젝트명: MenuExcelExporter
 * Version: 1.3 (URL 컬럼 위치 조정 및 파일명 패턴 통일)
 * 반영사항:
 * 1. [레이아웃] 'locaMenUrl'(URL) 컬럼을 순번 다음인 2번째로 이동
 * 2. [파일명] ApiExcelExporter와 동일한 패턴 적용: '메뉴목록_(yyyy-MM-dd_추출).xlsx' [cite: 2026-03-06]
 * 3. [기능 유지] 신규 LOCAmenu_new.json 규격 및 필수 5개 항목 추출 로직 유지 [cite: 2026-03-09]
 * 4. [에러 수정] JSON 내 비표준 주석 처리 가능하도록 ALLOW_COMMENTS 유지 [cite: 2026-03-09]
 * 5. [환경] config.properties UTF-8 로드 및 상세 로그 시스템 유지 [cite: 2026-03-09]
 */
public class MenuExcelExporter {

    private static String MENU_JSON_PATH = "";
    private static String MENU_OUTPUT_DIR = "";

    private static final ObjectMapper MAPPER = new ObjectMapper()
            .configure(JsonParser.Feature.ALLOW_COMMENTS, true);

    public static void main(String[] args) {
        loadConfig();

        if (MENU_JSON_PATH.isEmpty()) {
            System.err.println("[ERROR] MENU_JSON_PATH가 설정되지 않았습니다. config.properties를 확인하세요.");
            return;
        }

        long startTime = System.currentTimeMillis();
        // [v1.3] ApiExcelExporter와 동일한 날짜 형식 적용 [cite: 2026-03-06]
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd'_추출'"));

        System.out.println("===============================================================");
        System.out.println("[START] 메뉴 JSON 링크 추출 시작 (v1.3)");
        System.out.println("[INFO] 대상 파일: " + MENU_JSON_PATH);
        System.out.println("===============================================================");

        try {
            File jsonFile = new File(MENU_JSON_PATH);
            if (!jsonFile.exists()) {
                System.err.println("[ERROR] 파일을 찾을 수 없습니다.");
                return;
            }

            JsonNode root = MAPPER.readTree(jsonFile);
            List<MenuInfo> resultList = new ArrayList<>();

            if (root.isArray()) {
                traverseMenu(root, resultList);
            } else if (root.has("menu")) {
                traverseMenu(root.get("menu"), resultList);
            } else {
                traverseMenu(root, resultList);
            }

            saveToExcel(resultList, timestamp);

        } catch (Exception e) {
            System.err.println("[ERROR] 처리 중 오류: " + e.getMessage());
            e.printStackTrace();
        }

        System.out.println("\n[FINISH] 전체 작업 종료: " + (System.currentTimeMillis() - startTime) / 1000 + "초 소요");
    }

    private static void traverseMenu(JsonNode node, List<MenuInfo> list) {
        if (node.isArray()) {
            for (JsonNode item : node) {
                processItem(item, list);
            }
        } else {
            processItem(node, list);
        }
    }

    private static void processItem(JsonNode item, List<MenuInfo> list) {
        String url = item.path("locaMenUrl").asText("").trim();

        if (!url.isEmpty()) {
            MenuInfo info = new MenuInfo();
            info.locaMenId = item.path("locaMenId").asText("-");
            info.locaMenC = item.path("locaMenC").asText("-");
            info.locaMenIdNm = item.path("locaMenIdNm").asText("-");
            info.locaMenCNm = item.path("locaMenCNm").asText("-");
            info.locaMenSeaInfCn = item.path("locaMenSeaInfCn").asText("-");
            info.locaMenUrl = url;
            list.add(info);
            System.out.println("  > [수집] " + info.locaMenIdNm + " (" + info.locaMenUrl + ")");
        }

        if (item.has("sub") && item.get("sub").isArray()) {
            traverseMenu(item.get("sub"), list);
        }
    }

    private static void saveToExcel(List<MenuInfo> list, String ts) {
        // [v1.3] 파일명 패턴을 ApiExcelExporter와 동일하게 수정 [cite: 2026-03-06]
        String fileName = "메뉴목록_(" + ts + ").xlsx";
        File outFile = new File(MENU_OUTPUT_DIR, fileName);

        if (!outFile.getParentFile().exists()) outFile.getParentFile().mkdirs();

        try (Workbook wb = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(outFile)) {

            Sheet sheet = wb.createSheet("LOCA_Menu_Info");

            CellStyle headerStyle = wb.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            headerStyle.setBorderBottom(BorderStyle.THIN);
            Font font = wb.createFont(); font.setBold(true); headerStyle.setFont(font);

            // [v1.3] URL 컬럼을 2번째(인덱스 1)로 이동
            String[] headers = {"순번", "연결URL(locaMenUrl)", "메뉴ID(locaMenId)", "메뉴구분(locaMenC)", "메뉴명(locaMenIdNm)", "구분명(locaMenCNm)", "검색정보(locaMenSeaInfCn)"};
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                Cell c = headerRow.createCell(i);
                c.setCellValue(headers[i]);
                c.setCellStyle(headerStyle);
            }

            int rowIdx = 1;
            for (MenuInfo m : list) {
                Row r = sheet.createRow(rowIdx++);
                r.createCell(0).setCellValue(rowIdx - 1);
                r.createCell(1).setCellValue(m.locaMenUrl); // URL 위치 변경 [cite: 2026-03-09]
                r.createCell(2).setCellValue(m.locaMenId);
                r.createCell(3).setCellValue(m.locaMenC);
                r.createCell(4).setCellValue(m.locaMenIdNm);
                r.createCell(5).setCellValue(m.locaMenCNm);
                r.createCell(6).setCellValue(m.locaMenSeaInfCn);
            }

            for (int i = 0; i < headers.length; i++) {
                if (i == 1) sheet.setColumnWidth(i, 15000); // URL 너비 확장
                else if (i == 6) sheet.setColumnWidth(i, 12000);
                else sheet.setColumnWidth(i, 6000);
            }

            wb.write(fos);
            System.out.println("\n[SUCCESS] 엑셀 저장 완료: " + outFile.getAbsolutePath());
            System.out.println("[INFO] 총 추출 건수: " + list.size() + "건");

        } catch (Exception e) { e.printStackTrace(); }
    }

    private static void loadConfig() {
        Properties prop = new Properties();
        File configFile = new File("config.properties");
        if (configFile.exists()) {
            try (InputStreamReader isr = new InputStreamReader(new FileInputStream(configFile), StandardCharsets.UTF_8)) {
                prop.load(isr);
                MENU_JSON_PATH = prop.getProperty("MENU_JSON_PATH", "").trim();
                MENU_OUTPUT_DIR = prop.getProperty("MENU_OUTPUT_DIR", "").trim();
            } catch (IOException e) { e.printStackTrace(); }
        }
    }

    static class MenuInfo {
        String locaMenId, locaMenC, locaMenIdNm, locaMenCNm, locaMenSeaInfCn, locaMenUrl;
    }
}