import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.FillPatternType
import org.apache.poi.ss.usermodel.IndexedColors
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook

class ExcelStyles {

    static Workbook workbook
    static final String[] headers = ["Test Suite", "Test Case", "Status"]
    static FileOutputStream os
    static Sheet sheet
    static HSSFRow rowHeader


    static void main(String[] args) {


        def reportPath = "/Volumes/External/Development/2022/January/Groovy/GroovySnippets/files/report.xls"

        workbook = new HSSFWorkbook()
        os = new FileOutputStream(reportPath)
        sheet = workbook.createSheet("Default Sheet")
        rowHeader = sheet.createRow((short)0)
        createExcelTemplate()
        workbook.write(os)
        os.close()
        workbook.close()

    }


    def static createExcelTemplate(){
        headers.eachWithIndex{ String entry, int i ->
            def cell = rowHeader.createCell(i)
            if (entry == "Status"){
                println("Code is here")
                cell.setCellStyle(getSuccessCellType())
            }
            cell.setCellValue(entry)

        }
    }

    static CellStyle getSuccessCellType(){

        try{
            CellStyle style = workbook.createCellStyle()
            style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex())
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND)
            return style

        }catch(Exception e){
            throw new Exception(e.getMessage())
        }
    }
}
