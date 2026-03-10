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
 * Version: 1.0 (메뉴 JSON 링크 추출 도구)
 * 반영사항:
 * 1. [기능] JSON 내 'link'가 있는 항목을 재귀적으로 탐색하여 추출
 * 2. [데이터] menNm, note, link 정보를 엑셀로 변환
 * 3. [복구] config.properties UTF-8 로드 및 상세 로그 시스템 적용 [cite: 2026-03-09]
 */
public class MenuExcelExporter {

    private static String MENU_JSON_PATH = "";
    private static String MENU_OUTPUT_DIR = "";
    private static final ObjectMapper MAPPER = new ObjectMapper();

    public static void main(String[] args) {
        loadConfig();

        if (MENU_JSON_PATH.isEmpty()) {
            System.err.println("[ERROR] MENU_JSON_PATH가 설정되지 않았습니다.");
            return;
        }

        long startTime = System.currentTimeMillis();
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd'_추출'"));

        System.out.println("===============================================================");
        System.out.println("[START] 메뉴 JSON 링크 추출 시작 (v1.0)");
        System.out.println("[INFO] 대상 파일: " + MENU_JSON_PATH);
        System.out.println("===============================================================");

        try {
            File jsonFile = new File(MENU_JSON_PATH);
            JsonNode root = MAPPER.readTree(jsonFile);
            List<MenuInfo> resultList = new ArrayList<>();

            // "menu" 노드부터 탐색 시작
            if (root.has("menu")) {
                traverseMenu(root.get("menu"), resultList);
            }

            saveToExcel(resultList, timestamp);

        } catch (Exception e) {
            System.err.println("[ERROR] 처리 중 오류 발생: " + e.getMessage());
            e.printStackTrace();
        }

        System.out.println("\n[FINISH] 전체 작업 종료: " + (System.currentTimeMillis() - startTime) / 1000 + "초 소요");
    }

    /** [재귀 탐색] 하위 메뉴(sub)까지 모두 뒤져서 link가 있는 건만 수집합니다. */
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
        // link가 존재하고 비어있지 않은 경우 추출
        if (item.has("link") && !item.get("link").asText().isEmpty()) {
            MenuInfo info = new MenuInfo();
            info.menNm = item.path("menNm").asText("-");
            info.note = item.path("note").asText("-");
            info.link = item.path("link").asText();
            list.add(info);
            System.out.println("  > [발견] " + info.menNm + " (" + info.link + ")");
        }

        // 하위 메뉴(sub)가 있으면 다시 탐색 (재귀)
        if (item.has("sub") && item.get("sub").isArray() && item.get("sub").size() > 0) {
            traverseMenu(item.get("sub"), list);
        }
    }

    private static void saveToExcel(List<MenuInfo> list, String ts) {
        String fileName = "메뉴_링크_추출목록_(" + ts + ").xlsx";
        File outFile = new File(MENU_OUTPUT_DIR, fileName);

        try (Workbook wb = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(outFile)) {

            Sheet sheet = wb.createSheet("Menu_Link_Mapping");

            // 헤더 스타일 (Blue 테마 적용) [cite: 2026-03-09]
            CellStyle headerStyle = wb.createCellStyle();
            headerStyle.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
            headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerStyle.setAlignment(HorizontalAlignment.CENTER);
            headerStyle.setBorderBottom(BorderStyle.THIN);
            Font font = wb.createFont(); font.setBold(true); headerStyle.setFont(font);

            String[] headers = {"순번", "메뉴명(menNm)", "비고(note)", "연결 URL(link)"};
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
                r.createCell(1).setCellValue(m.menNm);
                r.createCell(2).setCellValue(m.note);
                r.createCell(3).setCellValue(m.link);
            }

            sheet.setColumnWidth(1, 8000);
            sheet.setColumnWidth(2, 10000);
            sheet.setColumnWidth(3, 15000);

            wb.write(fos);
            System.out.println("\n[SUCCESS] 엑셀 저장 완료: " + outFile.getAbsolutePath());
            System.out.println("[INFO] 총 추출 건수: " + list.size() + "건");

        } catch (Exception e) {
            System.err.println("[ERROR] 엑셀 생성 중 오류: " + e.getMessage());
        }
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
        String menNm, note, link;
    }
}