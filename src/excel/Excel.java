package excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Escuela
 */
public class Excel {
public Excel (File fileName) {
        List cellData=new ArrayList();
        try{
            FileInputStream fileInputStream=new FileInputStream(fileName);
            XSSFWorkbook workbook=new XSSFWorkbook(fileInputStream);
            XSSFSheet hssfSheet=workbook.getSheetAt(0);
            Iterator rowIterator=hssfSheet.rowIterator();
            while(rowIterator.hasNext()){
                XSSFRow hssfRow=(XSSFRow) rowIterator.next();
                Iterator iterator=hssfRow.cellIterator();
                List celltemp=new ArrayList();
                while(iterator.hasNext()){
                    XSSFCell hssfcell=( XSSFCell) iterator.next();
                    celltemp.add(hssfcell);
                }
                cellData.add(celltemp);
            }
        }catch(Exception e){
            e.printStackTrace();
        }
        obtener(cellData);
        marca(cellData);
    }
    private void obtener(List cellDataList){
        for (int i=0;i<cellDataList.size();i++){
            List celltempList=(List) cellDataList.get(i);
            for (int j=0;j<celltempList.size();j++){
                XSSFCell hssfCell=( XSSFCell) celltempList.get(j);
                String ValorCelda =hssfCell.toString();
                System.out.print(ValorCelda+"     ");
            }
             System.out.println();
        }
    }
    public void marca(List cellDataList){
    
    int posicion=0;
     Scanner LEER = new Scanner(System.in);
     System.out.print("inserte la posicion de la marca: ");
     posicion= LEER.nextInt();
    for (int i=posicion;i==posicion;i++){
            List celltempList=(List) cellDataList.get(posicion);
            for (int j=1;j==1;j++){
                XSSFCell hssfCell=( XSSFCell) celltempList.get(1);
                String ValorCelda =hssfCell.toString();
                EscribirEXCEL(ValorCelda);
            }
        }


}
    public static void EscribirEXCEL(String marca1) {
        String nombreArchivo = "Submarcas.xlsx";
        String hoja = "Hoja1";
        String submarca1 = "", color1 = "";
        XSSFWorkbook libro = new XSSFWorkbook();
        XSSFSheet hoja1 = libro.createSheet(hoja);
        Scanner teclado = new Scanner(System.in);
        System.out.println("Insertar submarca:");
        submarca1=(teclado.next());
        System.out.println("Insertar color:");
        color1=(teclado.next());
         
        // Cabecera de la hoja de excel
        String[] header = new String[]{"MARCA", "SUBMARCA", "COLOR"};

        // Contenido de la hoja de excel
        String[][] document = new String[][]{
            {marca1, submarca1, color1}
        };

        // Poner en negrita la cabecera
        CellStyle style = libro.createCellStyle();
        XSSFFont font = libro.createFont();
        font.setBold(true);
        style.setFont(font);

        // Generar los datos para el documento
        for (int i = 0; i <= document.length; i++) {
            XSSFRow row = hoja1.createRow(i);//se crea las filas
            for (int j = 0; j < header.length; j++) {
                if (i == 0) {//para la cabecera
                    XSSFCell cell = row.createCell(j);//se crea las celdas para la cabecera, junto con la posición
                    cell.setCellStyle(style); // se añade el style crea anteriormente 
                    cell.setCellValue(header[j]);//se añade el contenido

                }else {//para el contenido
                    XSSFCell cell = row.createCell(j);//se crea las celdas para la contenido, junto con la posición
                    cell.setCellValue(document[i - 1][j]); //se añade el contenido
                }
            }
        }

        // Crear el archivo
        try (OutputStream fileOut = new FileOutputStream(nombreArchivo)) {
            libro.write(fileOut);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    public static void main(String[] args) {
        File archivo1=new File("marcas.xlsx");
        if (archivo1.exists()){
            Excel obj=new Excel(archivo1);
        }
    }
    
}
