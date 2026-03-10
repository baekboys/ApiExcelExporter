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
 * Version: 1.2 (컬럼 순서 변경: URL을 맨 앞으로 배치)
 * 반영사항:
 * 1. [레이아웃] '연결 URL(link)' 컬럼을 0번 인덱스(맨 앞)로 이동하여 매핑 편의성 강화 [cite: 2026-03-09]
 * 2. [에러 수정] menu.json 내 비표준 주석 처리 가능하도록 ALLOW_COMMENTS 옵션 유지
 * 3. [기능 유지] JSON 내 'link'가 있는 항목을 재귀적으로 탐색하여 추출 [cite: 2026-03-09]
        * 4. [데이터] menNm, note, link 정보를 엑셀로 변환 [cite: 2026-03-09]
        * 5. [복구] config.properties UTF-8 로드 및 상세 로그 시스템 적용 [cite: 2026-03-09]
        */
public class MenuExcelExporter {

    private static String MENU_JSON_PATH = "";
    private static String MENU_OUTPUT_DIR = "";

    // [v1.1 수정] 주석이 포함된 JSON 파일을 읽을 수 있도록 ObjectMapper 설정 변경
    private static final ObjectMapper MAPPER = new ObjectMapper()
            .configure(JsonParser.Feature.ALLOW_COMMENTS, true);

    public static void main(String[] args) {
        loadConfig();

        if (MENU_JSON_PATH.isEmpty()) {
            System.err.println("[ERROR] MENU_JSON_PATH가 설정되지 않았습니다. config.properties를 확인하세요.");
            return;
        }

        long startTime = System.currentTimeMillis();
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy-MM-dd'_추출'"));

        System.out.println("===============================================================");
        System.out.println("[START] 메뉴 JSON 링크 추출 시작 (v1.2)");
        System.out.println("[INFO] 대상 파일: " + MENU_JSON_PATH);
        System.out.println("[INFO] 컬럼 순서: URL | 순번 | 메뉴명 | 비고");
        System.out.println("===============================================================");

        try {
            File jsonFile = new File(MENU_JSON_PATH);
            if (!jsonFile.exists()) {
                System.err.println("[ERROR] 파일을 찾을 수 없습니다: " + MENU_JSON_PATH);
                return;
            }

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
        if (item.has("link") && !item.get("link").asText().trim().isEmpty()) {
            MenuInfo info = new MenuInfo();
            info.menNm = item.path("menNm").asText("-");
            info.note = item.path("note").asText("-");
            info.link = item.path("link").asText();
            list.add(info);
            System.out.println("  > [Found] " + info.menNm + " (" + info.link + ")");
        }

        // 하위 메뉴(sub)가 있으면 다시 탐색 (재귀)
        if (item.has("sub") && item.get("sub").isArray() && item.get("sub").size() > 0) {
            traverseMenu(item.get("sub"), list);
        }
    }

    private static void saveToExcel(List<MenuInfo> list, String ts) {
        String fileName = "메뉴_링크_추출목록_(" + ts + ").xlsx";
        File outFile = new File(MENU_OUTPUT_DIR, fileName);

        if (!outFile.getParentFile().exists()) outFile.getParentFile().mkdirs();

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

            // [v1.2] URL을 맨 앞으로 이동 [cite: 2026-03-09]
            String[] headers = {"연결 URL(link)", "순번", "메뉴명(menNm)", "비고(note)"};
            Row headerRow = sheet.createRow(0);
            for (int i = 0; i < headers.length; i++) {
                Cell c = headerRow.createCell(i);
                c.setCellValue(headers[i]);
                c.setCellStyle(headerStyle);
            }

            int rowIdx = 1;
            for (MenuInfo m : list) {
                Row r = sheet.createRow(rowIdx++);
                // [v1.2] 데이터 매핑 순서 변경 [cite: 2026-03-09]
                r.createCell(0).setCellValue(m.link);    // 0: URL
                r.createCell(1).setCellValue(rowIdx - 1); // 1: 순번
                r.createCell(2).setCellValue(m.menNm);   // 2: 메뉴명
                r.createCell(3).setCellValue(m.note);    // 3: 비고
            }

            // [v1.2] 너비 설정 재조정 [cite: 2026-03-09]
            sheet.setColumnWidth(0, 15000); // URL
            sheet.setColumnWidth(1, 3000);  // 순번
            sheet.setColumnWidth(2, 8000);  // 메뉴명
            sheet.setColumnWidth(3, 10000); // 비고

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
            } catch (IOException e) {
                System.err.println("[ERROR] config.properties 로드 실패: " + e.getMessage());
            }
        }
    }

    static class MenuInfo {
        String menNm, note, link;
    }
}