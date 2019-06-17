/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package javaapplication1;

import java.util.Scanner;
import java.io.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;  
import java.util.ArrayList;
import java.util.Date; 
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.concurrent.TimeUnit;
import javax.swing.filechooser.FileSystemView;
import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFTableStyleInfo;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

/**
 *
 * @author nawue
 */

class PersFechaHora {
    public String nombre;
    public String fecha;
    public String horaEntrada;
    public String horaSalida;
    public Long minutosTotales;
    public String horaMinutoTotal;
}

class PersSemanal {
    public String nombre;
    public Long minutosSemanales;
}


public class JavaApplication1 {

    /**
     * @param args the command line arguments
     */
        
    public static ArrayList<PersFechaHora> datos = new ArrayList<PersFechaHora>();
    public static ArrayList<PersSemanal> datos2 = new ArrayList<PersSemanal>();
    
    public static void main(String[] args) throws FileNotFoundException, ParseException {
        obtenerDatos();
        crearExcel1();
    }
    
    public static void crearExcel1() {
        try {
            XSSFWorkbook workbook = new XSSFWorkbook();
            //nombre + fecha +seg
            XSSFSheet  sheet = workbook.createSheet("Registro");  

            XSSFRow rowhead = sheet.createRow((short)0);
            rowhead.createCell(0).setCellValue("Fecha");
            rowhead.createCell(1).setCellValue("Nombre");
            rowhead.createCell(2).setCellValue("Hora Entrada");
            rowhead.createCell(3).setCellValue("Hora Salida");
            rowhead.createCell(4).setCellValue("Horas Trabajadas");

            for(int i = 0; i < datos.size(); i++) {
                rowhead = sheet.createRow((short)(i+1));
                if (i != 0){
                    if (!datos.get(i-1).fecha.equals(datos.get(i).fecha))
                        rowhead.createCell(0).setCellValue(datos.get(i).fecha);
                } else {
                    rowhead.createCell(0).setCellValue(datos.get(i).fecha);
                }
                rowhead.createCell(1).setCellValue(datos.get(i).nombre);
                rowhead.createCell(2).setCellValue(datos.get(i).horaEntrada);
                rowhead.createCell(3).setCellValue(datos.get(i).horaSalida);
                rowhead.createCell(4).setCellValue(datos.get(i).horaMinutoTotal);
            }
            
            // Resize all columns to fit the content size
            for(int i = 0; i < 10; i++) {
                sheet.autoSizeColumn(i);
            }
            
            // Set which area the table should be placed in
            AreaReference reference = workbook.getCreationHelper().createAreaReference(
                    new CellReference(0, 0), new CellReference(datos.size(), 4));

            
            XSSFTable table = sheet.createTable(reference);
            table.setName("Test");
            table.setDisplayName("Test_Table");
            
            // For now, create the initial style in a low-level way
            table.getCTTable().addNewTableStyleInfo();
            table.getCTTable().getTableStyleInfo().setName("TableStyleMedium9");

            // Style the table
            XSSFTableStyleInfo style = (XSSFTableStyleInfo) table.getStyle();
            style.setName("TableStyleMedium9");
            style.setShowColumnStripes(false);
            style.setShowRowStripes(false);
            style.setFirstColumn(true);
            style.setLastColumn(false);
            style.setShowRowStripes(false);
            style.setShowColumnStripes(false);
            
            
            
            XSSFSheet  sheet2 = workbook.createSheet("RegistroSemanal");  
            XSSFCellStyle cellStyle = workbook.createCellStyle();
	    cellStyle.setFillForegroundColor(IndexedColors.LIGHT_BLUE.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            Font font = workbook.createFont(); 
            font.setColor(IndexedColors.WHITE.index); 
            font.setBold(true);
            cellStyle.setFont(font);

            
            XSSFRow rowhead2 = sheet2.createRow((short)0);
            Cell cell1 = rowhead2.createCell(0);
            cell1.setCellValue("Fecha:");
            cell1.setCellStyle(cellStyle);
            
            String auxFecha = datos.get(0).fecha + " -- " + datos.get(datos.size()-1).fecha;
            Cell cell2 = rowhead2.createCell(1);
            cell2.setCellValue(auxFecha);
            cell2.setCellStyle(cellStyle);

            rowhead2 = sheet2.createRow((short)2);
            rowhead2.createCell(0).setCellValue("Nombre");
            rowhead2.createCell(1).setCellValue("HorasTotalesSemanales");
            
            obtenerTrabajadores();
            
            for(int i = 0; i < datos2.size();i++) {
                rowhead2 = sheet2.createRow((short)(i+3));
                rowhead2.createCell(0).setCellValue(datos2.get(i).nombre);
                Long horasTotales = TimeUnit.MINUTES.toHours(datos2.get(i).minutosSemanales);
                Long minutosTotales = TimeUnit.MINUTES.toMinutes(datos2.get(i).minutosSemanales) - TimeUnit.HOURS.toMinutes(TimeUnit.MINUTES.toHours(datos2.get(i).minutosSemanales));
                String auxMinSem = "0";
                if (minutosTotales < 10) auxMinSem = horasTotales + ":0" + minutosTotales;
                else auxMinSem = horasTotales + ":" + minutosTotales;
                rowhead2.createCell(1).setCellValue(auxMinSem);
            }
            
            // Resize all columns to fit the content size
            for(int i = 0; i < 10; i++) {
                sheet2.autoSizeColumn(i);
            }
            
            // Set which area the table should be placed in
            AreaReference reference2 = workbook.getCreationHelper().createAreaReference(
                    new CellReference(2, 0), new CellReference(datos2.size() + 2, 1));

            
            XSSFTable table2 = sheet2.createTable(reference2);
            table2.setName("Test2");
            table2.setDisplayName("Test_Table2");
            
            // For now, create the initial style in a low-level way
            table2.getCTTable().addNewTableStyleInfo();
            table2.getCTTable().getTableStyleInfo().setName("TableStyleMedium9");

            // Style the table
            XSSFTableStyleInfo style2 = (XSSFTableStyleInfo) table2.getStyle();
            style2.setName("TableStyleMedium9");
            style2.setShowColumnStripes(false);
            style2.setShowRowStripes(false);
            style2.setFirstColumn(true);
            style2.setLastColumn(false);
            style2.setShowRowStripes(false);
            style2.setShowColumnStripes(false);

            
            int num = 1;
            File home = FileSystemView.getFileSystemView().getHomeDirectory(); 
            String absPath = home.getAbsolutePath();

            File archivo = new File(absPath + "/Registro" + num + ".xlsx");
            while (archivo.exists()) {
                num++;
                archivo = new File(absPath + "/Registro" + num + ".xlsx");
            }
            FileOutputStream fileOut = new FileOutputStream(absPath + "/Registro" + num + ".xlsx");
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
            System.out.println("Your excel file has been generated!");                  
            

        } catch ( Exception ex ) {
            System.out.println(ex);
        }
    }
    
    
    public static void obtenerDatos() throws FileNotFoundException, ParseException{
        FileReader inputFile = new FileReader("001_GLog.txt");
        Scanner parser = new Scanner(inputFile);
        parser.nextLine();
        
        while (parser.hasNextLine())
        {
            
            String line = parser.nextLine();
            String[] att = line.split("\t");
            PersFechaHora pfh = new PersFechaHora();
            pfh.nombre = att[3];
            
            String aux = att[6].split("  ")[0];
            SimpleDateFormat auxSDF = new SimpleDateFormat("yyyy/MM/dd");
            Date auxDate = auxSDF.parse(aux);
            pfh.fecha = auxSDF.format(auxDate);
            
            /*aux = att[6].split("  ")[1];
            SimpleDateFormat auxSDF2 = new SimpleDateFormat("hh:mm:ss");
            Date auxHora = auxSDF2.parse(aux);*/
            pfh.horaEntrada = att[6].split("  ")[1];
            pfh.horaSalida = buscaHoraSalida(pfh.nombre, pfh.fecha, pfh.horaEntrada);       
            if (!pfh.horaSalida.equals("no hay registro de salida")) {
                Date hora1 = new SimpleDateFormat("hh:mm:ss").parse(pfh.horaEntrada);
                Date hora2 = new SimpleDateFormat("hh:mm:ss").parse(pfh.horaSalida);
                pfh.minutosTotales = (hora2.getTime() - hora1.getTime())/(1000*60);
                Long horasTotales = TimeUnit.MINUTES.toHours(pfh.minutosTotales);
                Long minutosTotales = TimeUnit.MINUTES.toMinutes(pfh.minutosTotales) - TimeUnit.HOURS.toMinutes(TimeUnit.MINUTES.toHours(pfh.minutosTotales));
                if (minutosTotales < 10) pfh.horaMinutoTotal = horasTotales + ":0" + minutosTotales;
                else pfh.horaMinutoTotal = horasTotales + ":" + minutosTotales;
            } else {
               pfh.minutosTotales = new Long(0);
               pfh.horaMinutoTotal = "0";
            }
            if (!comprobarExiste(pfh.nombre, pfh.fecha)) datos.add(pfh);
            
        }
    }
    
    public static String buscaHoraSalida(String nombre, String fecha, String horaEntrada) throws FileNotFoundException {
        FileReader inputFile = new FileReader("001_GLog.txt");
        Scanner parser = new Scanner(inputFile);
        parser.nextLine();
        
        while (parser.hasNextLine())
        {
            String line = parser.nextLine();
            String[] att = line.split("\t");
            String aux1 = att[6].split("  ")[0];
            String aux2 = att[6].split("  ")[1];
            if (att[3].equals(nombre) &&  aux1.equals(fecha) && !aux2.equals(horaEntrada)) return aux2;
        }
        
        return "no hay registro de salida";
    }
        
    public static Boolean comprobarExiste(String nombre, String fecha) {
        for(int i = 0; i < datos.size(); i++) {
            if(datos.get(i).nombre.equals(nombre) && datos.get(i).fecha.equals(fecha)) return true;
        }
        return false;
    }
    
    public static void obtenerTrabajadores(){
        ArrayList<String> trabajadores = new ArrayList<>();
        
        for(int i = 0; i < datos.size(); i++) {
            if(!trabajadores.contains(datos.get(i).nombre))
                trabajadores.add(datos.get(i).nombre);
        }
       
        for(int i = 0; i < trabajadores.size(); i++) {
            PersSemanal ps = new PersSemanal();
            ps.nombre = trabajadores.get(i);
            ps.minutosSemanales = obtenerMinutos(trabajadores.get(i));
            datos2.add(ps);
        }
    }
    
    public static Long obtenerMinutos(String nombre) {
        Long minutos = new Long(0);
        for(int i = 0; i < datos.size(); i++) {
            if (datos.get(i).nombre.equals(nombre))
                minutos += datos.get(i).minutosTotales;
        }
        return minutos;
    }
}
