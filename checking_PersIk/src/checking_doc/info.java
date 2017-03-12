/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package checking_doc;

import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;
import javax.swing.JOptionPane;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

/**
 *
 * @author Maxim
 * Класс, в котором производится анализ документа с выдачей уведомлений об ошибках
 */
public class info {
    public String wave_to_doc;
    private static String HighKolontit;
    private static String LowKolontit;
    
    private static void getKolontitules(String wave) {
        try{
            FileInputStream fileInputStream = new FileInputStream(wave);
            // открываем файл и считываем его содержимое в объект XWPFDocument
            XWPFDocument docxFile = new XWPFDocument(OPCPackage.open(fileInputStream));
            XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(docxFile);
 
            // считываем верхний колонтитул (херед документа)
            XWPFHeader docHeader = headerFooterPolicy.getDefaultHeader();
            HighKolontit = docHeader.getText();
            XWPFFooter docFooter = headerFooterPolicy.getDefaultFooter();
            LowKolontit = docFooter.getText();
            fileInputStream.close();
        }
        catch(Exception e){
            
        }
    }
    public void messages(){
        getKolontitules(wave_to_doc);
        boolean error = true;
        if(HighKolontit==null){
            JOptionPane.showMessageDialog(null,"<html><h2>Внимание!!!</h2><i>Верхний колонтитул пуст </i>");
            error = false;
        }
        if(LowKolontit==null){
            JOptionPane.showMessageDialog(null,"<html><h2>Внимание!!!</h2><i>Нижний колонтитул пуст </i>");
            error = false;
        }
        try{
            FileInputStream fileInputStream = new FileInputStream(wave_to_doc);
            XWPFDocument docxFile = new XWPFDocument(OPCPackage.open(fileInputStream));
            List<XWPFParagraph> paragraphs = docxFile.getParagraphs();
            String UDK;
            UDK = paragraphs.get(0).getText();
            if(UDK.indexOf("УДК")==-1){
                JOptionPane.showMessageDialog(null,"<html><h2>Внимание!!!</h2><i>Отсутствует строка УДК </i>");
                error = false;
            }
            Boolean cursive_authors;
            List <XWPFRun>  run = paragraphs.get(2).getRuns();
            cursive_authors = run.get(0).isItalic();
            if(!cursive_authors){
                JOptionPane.showMessageDialog(null,"<html><h2>Внимание!!!</h2><i>Авторы не выделены курсивом</i>");
                error = false;
            }
            Boolean global_text;
            List <XWPFRun>  run1 = paragraphs.get(6).getRuns();
            global_text = run1.get(0).isItalic();
            if(global_text){
                JOptionPane.showMessageDialog(null,"<html><h2>Внимание!!!</h2><i>Основной текст не должен быть курсивным</i>");
                error = false;
            }
            if(UDK.equals("УДК")||UDK.equals("УДК ")){
                JOptionPane.showMessageDialog(null,"<html><h2>Внимание!!!</h2><i>Отсутствует номер УДК </i>");
                error = false;
            }

            
            
            
            if (error){
                JOptionPane.showMessageDialog(null,"<html><h2>Удивительно!!!</h2><i>Похоже, будто все правильно, <br>но все-таки проверь</i>");
            }
            if (!error){
                JOptionPane.showMessageDialog(null,"<html><h2>Файл отформатирован неправильно</h2><i>Исправьте в соответствии с полученными замечаниями</i>");
            }
        }
        catch(Exception e){
            JOptionPane.showMessageDialog(null, e);
        }
    }
}
