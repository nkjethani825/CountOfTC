import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Color;

import java.io.*;

public class CountOfTCs {

    public  static  void main(String args[]) throws IOException {
        File folder=new File("C:\\Users\\MP\\Downloads\\CountOfTC\\src\\test\\java");
        File[] files=folder.listFiles();
        int count;
        String filename = "C:\\Users\\MP\\Downloads\\NewExcelFile1.xls" ;
        HSSFWorkbook workbook = new HSSFWorkbook();

        HSSFSheet sheet = workbook.createSheet("Count");

        HSSFCellStyle style = workbook.createCellStyle();
        style.setBorderTop(BorderStyle.DOUBLE);
        style.setBorderBottom(BorderStyle.THIN);
        //  style.setFillBackgroundColor(HSSFColor.toHSSFColor(new Color())

        // We also define the font that we are going to use for displaying the
        // data of the cell. We set the font to ARIAL with 20pt in size and
        // make it BOLD and give blue as the color.
        HSSFFont font = workbook.createFont();
        font.setFontName(HSSFFont.FONT_ARIAL);
        font.setFontHeightInPoints((short) 15);

        font.setBold(true);
        //  font.setColor(HSSFColor.HSSFColorPredefined.BLUE);
        style.setFont(font);
        HSSFRow rowhead = sheet.createRow((short)0);
        HSSFCell head1=rowhead.createCell(0);
        head1.setCellStyle(style);
        head1.setCellValue(new HSSFRichTextString("Sr. No."));
        HSSFCell head2=rowhead.createCell(1);
        head2.setCellStyle(style);
        head2.setCellValue(new HSSFRichTextString("Module Name"));
        HSSFCell head3=rowhead.createCell(2);
        head3.setCellStyle(style);
        head3.setCellValue(new HSSFRichTextString("TC Count"));
        for(int i=0;i<files.length;i++)
        {
            File file=files[i];
            if(file.getName().contains("Test")) {

                HSSFRow row = sheet.createRow((short)i+1);
                HSSFCell cell1=row.createCell(0);
                cell1.setCellValue(new HSSFRichTextString(new Integer(i+1).toString()));
                HSSFCell cell2=row.createCell(1);
                cell2.setCellValue(new HSSFRichTextString(file.getName()));
                HSSFCell cell3=row.createCell(2);
                count=countWord("@Test",file.getAbsolutePath());
                cell3.setCellValue(new HSSFRichTextString(new Integer(count).toString()));
            }
        }

        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(2);
        FileOutputStream fileOut = new FileOutputStream(filename);
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
        System.out.println("Your excel file has been generated!");
    }

    public static int countWord(String word,String filename) throws IOException {
        FileInputStream fis=new FileInputStream(filename);
        DataInputStream dis=new DataInputStream(fis);
        BufferedReader br=new BufferedReader(new InputStreamReader(dis));
        int count=0;
        String line;
        boolean flag=false;
        while((line=br.readLine())!=null)
        {
            if(flag==true) {
                System.out.println(line);
                flag=false;
            }
            if(line.contains("/*"))
            {
                line=br.readLine();
                while (line.contains("*/"))
                   line= br.readLine();
            }
            String[] wordsOfLine=line.split(" ");
            for (String words:wordsOfLine)
            {
                if(words.contains("//"))
                    break;

                if(words.equals(word)) {
                    count++;
                    flag=true;
                    break;
                }
            }
        }

        br.close();

        return  count;

    }
}
