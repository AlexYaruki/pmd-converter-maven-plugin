package com.alexyaruki.sca.pmd;

import net.sourceforge.pmd.RulePriority;
import org.apache.maven.plugin.AbstractMojo;
import org.apache.maven.plugin.MojoExecutionException;
import org.apache.maven.plugin.MojoFailureException;
import org.apache.maven.plugins.annotations.Mojo;
import org.apache.maven.plugins.annotations.Parameter;
import org.apache.maven.project.MavenProject;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;



@Mojo(name = "convert")
public class ExcelConverter extends AbstractMojo {

    private static class Counter {

        private int value;

        Counter(int value) {
            this.value = value;
        }

        public void increment() {
            value++;
        }

        public void decrement() {
            value--;
        }

        public int getValue() {
            return value;
        }

    }

    public static void main(String[] args) throws ParserConfigurationException, IOException, SAXException {
        Path pmdPath = Paths.get("pmd.xml");
        processPMD(pmdPath, pmdPath);
    }

    /**
     * Object representing current Maven project.
     */
    @Parameter(defaultValue = "${project}", readonly = true, required = true)
    private MavenProject project;

    private static void processPMD(Path pmdPath, Path pmdExcelPath) throws ParserConfigurationException, SAXException, IOException {
        if (!pmdPath.toFile().exists()) {
            return;
        }

        DocumentBuilder documentBuilder = DocumentBuilderFactory.newInstance().newDocumentBuilder();
        Document document = documentBuilder.parse(pmdPath.toFile());
        Element root = document.getDocumentElement();
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet pmdViolations = workbook.createSheet("PMD Violations");
        createHeader(pmdViolations);

        Counter counter = new Counter(1);
        NodeList files = root.getElementsByTagName("file");
        for (int i = 0; i < files.getLength(); i++) {
            Node file = files.item(i);
            processFile(pmdViolations, (Element) file, counter);
        }
        pmdViolations.autoSizeColumn(0);
        pmdViolations.autoSizeColumn(1);
        pmdViolations.autoSizeColumn(2);
        pmdViolations.autoSizeColumn(3);
        pmdViolations.autoSizeColumn(4);
        pmdViolations.autoSizeColumn(5);
        pmdViolations.autoSizeColumn(6);
        pmdViolations.autoSizeColumn(7);

        try (FileOutputStream outputStream = new FileOutputStream(pmdExcelPath.toFile())) {
            workbook.write(outputStream);
        }
    }

    private static void processFile(XSSFSheet pmdViolations, Element file, Counter counter) {
        String filePath = file.getAttribute("name");
        filePath = filePath.substring(filePath.lastIndexOf("main/java") + 10);
        NodeList violations = file.getElementsByTagName("violation");
        for (int i = 0; i < violations.getLength(); i++) {
            Element violation = (Element) violations.item(i);
            processViolation(pmdViolations, filePath, violation, counter);
        }
    }

    private static void processViolation(XSSFSheet pmdViolations, String filePath, Element violation, Counter counter) {
        XSSFRow violationRow = pmdViolations.createRow(counter.getValue());
        counter.increment();
        violationRow.createCell(0).setCellValue(filePath);
        String beginLine = violation.getAttribute("beginline");
        String endLine = violation.getAttribute("endline");
        String beginColumn = violation.getAttribute("begincolumn");
        String endColumn = violation.getAttribute("endcolumn");
        violationRow.createCell(1).setCellValue(beginLine);
        violationRow.createCell(2).setCellValue(endLine);
        violationRow.createCell(3).setCellValue(beginColumn);
        violationRow.createCell(4).setCellValue(endColumn);
        String description = violation.getTextContent();
        description = description.replace("\n","").replaceAll("( +)"," ").trim();
        RulePriority priority = RulePriority.valueOf(Integer.parseInt(violation.getAttribute("priority")));
        violationRow.createCell(5).setCellValue(violation.getAttribute("ruleset") + " -> " + violation.getAttribute("rule"));
        violationRow.createCell(6).setCellValue(priority.getName());
        violationRow.createCell(7).setCellValue(description);
        violationRow.setHeight((short) -1);
    }

    private static void createHeader(XSSFSheet pmdViolations) {
        XSSFRow headerRow = pmdViolations.createRow(0);
        headerRow.createCell(0).setCellValue("Source file");
        headerRow.createCell(1).setCellValue("Begin line");
        headerRow.createCell(2).setCellValue("End line");
        headerRow.createCell(3).setCellValue("Begin column");
        headerRow.createCell(4).setCellValue("End column");
        headerRow.createCell(5).setCellValue("Category");
        headerRow.createCell(6).setCellValue("Priority");
        headerRow.createCell(7).setCellValue("Description");
    }

    @Override
    public void execute() throws MojoExecutionException, MojoFailureException {
        Path pmdPath = Paths.get(project.getBuild().getDirectory(), "pmd.xml");
        Path pmdExcelPath = Paths.get(project.getBuild().getDirectory(),"pmd.xlsx");
        if(!pmdPath.toFile().exists()) {
            throw new MojoExecutionException("pmd.xml not found in " + project.getBuild().getDirectory() + " directory");
        }
        try {
            processPMD(pmdPath,pmdExcelPath);
            getLog().info("pmd.xml converted to pmd.xslx");
        } catch (ParserConfigurationException | SAXException | IOException e) {
            throw new MojoFailureException("Cannot convert PMD file",e);
        }
    }
}
