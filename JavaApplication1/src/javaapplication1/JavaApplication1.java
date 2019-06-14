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
import org.apache.poi.hssf.usermodel.HSSFRow;
import  org.apache.poi.hssf.usermodel.HSSFSheet;
import  org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 *
 * @author nawue
 */

class PersFechaHora {
    public String nombre;
    public String fecha;
    public String horaEntrada;
    public String horaSalida;
}


public class JavaApplication1 {

    /**
     * @param args the command line arguments
     */
        
    public static ArrayList<PersFechaHora> datos = new ArrayList<PersFechaHora>();
    
    public static void main(String[] args) throws FileNotFoundException, ParseException {
        obtenerDatos();
        crearExcel();
    }
    
    public static void crearExcel() {
        String filename = "registro.xls";
        try {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("Registro");  

            HSSFRow rowhead = sheet.createRow((short)0);
            rowhead.createCell(0).setCellValue("Fecha");
            rowhead.createCell(1).setCellValue("Nombre");
            rowhead.createCell(2).setCellValue("Hora Entrada");
            rowhead.createCell(3).setCellValue("Hora Salida");
            rowhead.createCell(4).setCellValue("Horas Trabajadas");
            
            for(int i = 0; i < datos.size(); i++) {
                rowhead = sheet.createRow((short)(i+1));
                rowhead.createCell(0).setCellValue(datos.get(i).fecha);
                rowhead.createCell(1).setCellValue(datos.get(i).nombre);
                rowhead.createCell(2).setCellValue(datos.get(i).horaEntrada);
                rowhead.createCell(3).setCellValue(datos.get(i).horaSalida);
                if (!datos.get(i).horaSalida.equals("no hay registro de salida")) {
                    Date hora1 = new SimpleDateFormat("hh:mm:ss").parse(datos.get(i).horaEntrada);
                    Date hora2 = new SimpleDateFormat("hh:mm:ss").parse(datos.get(i).horaSalida);
                    long hora11 = hora1.getHours();
                    long hora22 = hora2.getHours();
                    rowhead.createCell(4).setCellValue(hora22 - hora11);
                } else {
                    rowhead.createCell(4).setCellValue("no se puede calcular");
                }
            }
            
            FileOutputStream fileOut = new FileOutputStream(filename);
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
            System.out.println("wecw1: " +  pfh.horaEntrada + " 2: " + att[6].split("  ")[1]);
            pfh.horaSalida = buscaHoraSalida(pfh.nombre, pfh.fecha, pfh.horaEntrada);
            
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
            System.out.println("1: " + horaEntrada + "2: " + aux2);
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
}
