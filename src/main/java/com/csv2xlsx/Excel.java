package com.csv2xlsx;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.LineNumberReader;
import java.io.PrintStream;
import java.util.Date;
import java.util.Properties;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel
{
  private static XSSFWorkbook wb;
  
  public Excel() {}
  
  public static Properties GetProp() throws IOException
  {
    BufferedReader br = null;
    
    Properties config = new Properties();
    

    File diretorioBase = new File("src/converterCSVtoXLSX.properties");
    
    String file = getAppFilePath(diretorioBase.getParentFile().getAbsolutePath() + 
      File.separatorChar + File.separatorChar + 
      "converterCSVtoXLSX.properties");
    br = new BufferedReader(new FileReader(file));
    config.load(br);
    return config;
  }
  
  public static String getAppFilePath(String fileName)
  {
    File file = new File(fileName);
    if (file.exists()) {
      return fileName;
    }
    

    String path = System.getenv("PATH");
    if (path == null) {
      path = System.getenv("Path");
    }
    
    if (path != null) {
      String separator = System.getProperty("path.separator");
      String[] paths = path.split(separator);
      for (int i = 0; i < paths.length; i++) {
        String appFilePath = paths[i] + File.separatorChar + fileName;
        file = new File(appFilePath);
        if (file.exists()) {
          return appFilePath;
        }
      }
    }
    


    return fileName;
  }
  
  public static void main(String[] args)
    throws IOException
  {
    boolean help = false;
    boolean bradesco = false;
    boolean duas_casas = false;
    

    Properties cfg = GetProp();
    


    if ((args.length > 5) || (args.length < 5)) {
      help = true;
    }
    if (help) {
      System.out.println("PARAMETROS INVALIDOS, FAVOR VERIFICAR CONFORME ABAIXO");
      System.out.println("PARAMENTRO 1 dentro de Aspas = Entrar com o arquivo de entrada no formtado CSV");
      System.out.println("PARAMENTRO 2 dentro de Aspas = Nome do arquivo de saida com a extensao em xlsx OBRIGATÓRIA.");
      System.out.println("PARAMENTRO 3 dentro de Aspas = Titulo do relatorio");
      System.out.println("PARAMENTRO 4 Fora de Aspas = Qtd limite de linhas que o programa deve considerar para inicar a conversão");
      System.out.println("PARAMENTRO 5 dentro de Aspas = SIM/NAO caso trate a formatacao para numero quando a coluna for VALOR ou valor");
      
      System.exit(0);
    }
    
    String ArquivoEntrada = "";
    ArquivoEntrada = args[0];
    
    String ArquivoSainda = "";
    ArquivoSainda = args[1];
    
    String Titulo = "";
    Titulo = args[2];
    
    int limiteLinhas = Integer.parseInt(args[3]);
    

    String TratarNumeros = "";
    TratarNumeros = args[4];
    
    if (TratarNumeros.equalsIgnoreCase("SIM"))
    {
      bradesco = true;
    }
    if (TratarNumeros.equalsIgnoreCase("NAO"))
    {
      bradesco = false;
    }
    
    Date data = new Date();
    
    System.out.println(data);
    System.out.println("Inicio do processamento da conversão do arquivo: " + ArquivoEntrada + " Para o formato xlsx");
    

    FileOutputStream fileOutPut = new FileOutputStream(ArquivoSainda);
    


    wb = new XSSFWorkbook();
    
    Sheet sheet = wb.createSheet();
    
    sheet.setDisplayGridlines(false);
    
    Row row = sheet.createRow(0);
    
    Cell cell = row.createCell(0);
    
    CellStyle estilo = wb.createCellStyle();
    
    String arquivoCSV = ArquivoEntrada;
    BufferedReader br = null;
    String linha = "";
    String csvDivisor = ";";
    boolean campoNumber = false;
    


    try
    {
      br = new BufferedReader(new FileReader(arquivoCSV));
      String controle = br.readLine();
      String[] colunas = controle.split(csvDivisor, -1);
      int QtdColunas = colunas.length - 1;
      
      File arquivoLeitura = new File(arquivoCSV);
      LineNumberReader linhaLeitura = new LineNumberReader(new FileReader(arquivoLeitura));
      linhaLeitura.skip(arquivoLeitura.length());
      int qtdLinha = linhaLeitura.getLineNumber() + 1;
      
      if (limiteLinhas < qtdLinha) {
        System.out.println("Processo abortado!");
        System.out.println("Quantidade de linha: " + qtdLinha + " menor que o parametro: " + limiteLinhas);
        System.exit(0);
      }
      linhaLeitura.close();
      
      Font font = wb.createFont();
      font.setFontHeightInPoints((short)10);
      estilo = null;
      estilo = wb.createCellStyle();
      estilo.setBorderBottom((short)1);
      estilo.setBorderLeft((short)1);
      estilo.setBorderRight((short)1);
      
      for (int i = 0; i <= QtdColunas; i++)
      {
        int cont = 6;
        br = new BufferedReader(new FileReader(arquivoCSV));
        
        while ((linha = br.readLine()) != null) {
          String[] pais = linha.split(csvDivisor, -1);
          
          if (i == 0)
          {
            row = sheet.createRow(cont);
            cell = row.createCell(i);
            cell.setCellValue(pais[i]);
            
            estilo.setFont(font);
            
            cell.setCellStyle(estilo);
          } else {
            String num_duas_casas = cfg.getProperty("INDICE_COLUNA_FORMAT_NUMBER");
            String[] a = num_duas_casas.split(",");
            

            if (bradesco)
            {
              int Qtd = a.length;
              for (int f = 0; f < Qtd; f++)
              {
                int aux = Integer.parseInt(a[f]);
                if (i == aux) {
                  duas_casas = true;
                }
              }
            }
            if ((bradesco) && (cont == 6) && (duas_casas)) {
              campoNumber = true;
            }
            



            if ((campoNumber) && (cont >= 7))
            {
              String value = pais[i];
              value = value.replaceAll(",", ".");
              Double Valor = Double.valueOf(Double.parseDouble(value));
              row = sheet.getRow(cont);
              cell = row.createCell(i);
              cell.setCellValue(new Double(Valor.doubleValue()).doubleValue());
              
              estilo.setFont(font);
              
              cell.setCellStyle(estilo);
              duas_casas = false;
            }
            else {
              row = sheet.getRow(cont);
              cell = row.createCell(i);
              cell.setCellValue(pais[i]);
              
              estilo.setFont(font);
              
              cell.setCellStyle(estilo);
            }
          }
          
          cont++;
        }
        campoNumber = false;
        br.close();
      }
      

      for (int i = 0; i <= QtdColunas; i++) {
        row = sheet.getRow(6);
        cell = row.getCell((short)i);
        font = wb.createFont();
        font.setFontHeightInPoints((short)11);
        font.setBoldweight((short)700);
        font.setColor(IndexedColors.WHITE.getIndex());
        estilo = row.getRowStyle();
        estilo = null;
        estilo = wb.createCellStyle();
        
        estilo.setFillForegroundColor(IndexedColors.RED.getIndex());
        estilo.setFillPattern((short)1);
        estilo.setBorderBottom((short)1);
        estilo.setBorderLeft((short)1);
        estilo.setBorderRight((short)1);
        estilo.setFont(font);
        cell.setCellStyle(estilo);
      }
      

      font = wb.createFont();
      font.setBoldweight((short)700);
      font.setFontHeightInPoints((short)15);
      
      CellStyle borderStyle = wb.createCellStyle();
      borderStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
      borderStyle.setFillPattern((short)1);
      borderStyle.setBorderBottom((short)1);
      borderStyle.setBorderLeft((short)1);
      borderStyle.setBorderRight((short)1);
      borderStyle.setBorderTop((short)1);
      borderStyle.setAlignment((short)2);
      row = sheet.createRow(5);
      for (int i = 0; i <= QtdColunas; i++) {
        cell = row.createCell((short)i);
        if (i == 0) {
          cell.setCellValue(Titulo);
        }
        
        cell.setCellStyle(borderStyle);
        borderStyle.setFont(font);
      }
      

      sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(5, 5, 0, QtdColunas));
      
      row = null;
      cell = null;
      row = sheet.createRow(0);
      
      File diretorio = new File("src/logo-bradesco.jpg");
      
      FileInputStream InputStream = new FileInputStream(diretorio);
      byte[] bytes = IOUtils.toByteArray(InputStream);
      cell = row.createCell(0);
      
      int pictureIdx = wb.addPicture(bytes, 5);
      
      InputStream.close();
      CreationHelper helper = wb.getCreationHelper();
      Drawing drawing = sheet.createDrawingPatriarch();
      ClientAnchor anchor = helper.createClientAnchor();
      
      anchor.setCol1(0);
      anchor.setRow1(0);
      anchor.setCol2(9);
      anchor.setRow2(3);
      
      Picture pict = drawing.createPicture(anchor, pictureIdx);
      
      wb.write(fileOutPut);
      fileOutPut.close();
    }
    catch (FileNotFoundException e) {
      e.printStackTrace();
      


      if (br != null) {
        try {
          br.close();
          fileOutPut.close();
        }
        catch (IOException e1) {
          e1.printStackTrace();
          System.out.println("Processo finalizado com ERRO!");
        }
      }
    }
    catch (IOException e)
    {
      e.printStackTrace();
      
      if (br != null) {
        try {
          br.close();
          fileOutPut.close();
        }
        catch (IOException e2) {
          e2.printStackTrace();
          System.out.println("Processo finalizado com ERRO!");
        }
      }
    }
    finally
    {
      if (br != null) {
        try {
          br.close();
          fileOutPut.close();
        }
        catch (IOException e) {
          e.printStackTrace();
          System.out.println("Processo finalizado com ERRO!");
        }
      }
    }
    
    System.out.println("Arquivo gerado com sucesso: " + ArquivoSainda);
    System.out.println("Processo finalizado com Sucesso!");
  }
}
