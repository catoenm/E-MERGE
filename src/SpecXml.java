
import java.awt.Toolkit;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import jxl.Cell;
import jxl.CellType;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author Mitchell
 */
public class GUI extends javax.swing.JFrame {

    /**
     * Creates new form GUI
     */
    public GUI() {
        initComponents();
    }
    
    //VARIABLE DECLARATION
    static File specSheet;
    static File fXmlFile;
    String modelNumber;
    String screwDiameter;
    boolean again = true;
		int modelColomnNumber = -1, colomnNumber = -1;
		ArrayList <String> specSheetValues = new ArrayList<String>();
		ArrayList <String> specSheetVariableNames = new ArrayList<String>();
		ArrayList <String> xmlValues = new ArrayList<String>();
		ArrayList <String> xmlVariableNames = new ArrayList<String>();
                ArrayList <String> mistakes = new ArrayList<String>();
    
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jLabel1 = new javax.swing.JLabel();
        jScrollPane1 = new javax.swing.JScrollPane();
        area = new javax.swing.JTextArea();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        StartButton = new java.awt.Button();
        jLabel4 = new javax.swing.JLabel();
        button1 = new java.awt.Button();
        button2 = new java.awt.Button();
        specSheetPath = new javax.swing.JTextField();
        machineDataPath = new javax.swing.JTextField();
        proceedButton = new java.awt.Button();
        modelNumberText = new javax.swing.JComboBox();
        screwDiameterText = new javax.swing.JComboBox();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jLabel1.setFont(new java.awt.Font("Copperplate Gothic Bold", 0, 48)); // NOI18N
        jLabel1.setText("MACHINE DATA COMPARE");

        area.setColumns(20);
        area.setRows(5);
        jScrollPane1.setViewportView(area);

        jLabel2.setFont(new java.awt.Font("Dialog", 0, 22)); // NOI18N
        jLabel2.setText("Model Number:");

        jLabel3.setFont(new java.awt.Font("Dialog", 0, 20)); // NOI18N
        jLabel3.setText("Screw Diameter:");

        StartButton.setActionCommand("Start");
        StartButton.setLabel("Start");
        StartButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                StartButtonActionPerformed(evt);
            }
        });

        jLabel4.setFont(new java.awt.Font("Dialog", 0, 24)); // NOI18N
        jLabel4.setText("Select Spec Sheet and XML file:");

        button1.setLabel("Select Spec Sheet");
        button1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                button1ActionPerformed(evt);
            }
        });

        button2.setLabel("Select Machine Data");
        button2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                button2ActionPerformed(evt);
            }
        });

        specSheetPath.setEditable(false);

        machineDataPath.setEditable(false);

        proceedButton.setActionCommand("Proceed");
        proceedButton.setLabel("Open Machine Data Modification");
        proceedButton.setName(""); // NOI18N
        proceedButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                proceedButtonActionPerformed(evt);
            }
        });

        modelNumberText.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "EM1-15", "EM1-15HS", "EM1-30", "EM1-30HS", "EM1-30UHS", "EM1-30-22", "EM2-50", "EM2-50HS", "EM2-50UHS", "EM2-80", "EM2-80HS", "EM2-80UHS", "EM3-100", "EM3-100HS", "EM3-200", "EM3-200HS", "EM3-250", "EM4-350", "EM4-550", "EM4-350HP", "EM4-550HP", "EM4-350HS", "EM4-550MHS", "EM4-550HS" }));

        screwDiameterText.setModel(new javax.swing.DefaultComboBoxModel(new String[] { "12", "14", "16", "18", "20", "22", "25", "28", "32", "35", "38", "40", "45", "50", "55" }));

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jScrollPane1)
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(layout.createSequentialGroup()
                                .addComponent(jLabel4)
                                .addGap(100, 100, 100))
                            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, layout.createSequentialGroup()
                                .addComponent(jLabel2)
                                .addGap(18, 18, 18)
                                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                                    .addComponent(modelNumberText, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                    .addComponent(specSheetPath, javax.swing.GroupLayout.DEFAULT_SIZE, 280, Short.MAX_VALUE))))
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addGroup(layout.createSequentialGroup()
                                .addGap(10, 10, 10)
                                .addComponent(button2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                            .addGroup(layout.createSequentialGroup()
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                                .addComponent(jLabel3)))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addComponent(machineDataPath)
                            .addComponent(screwDiameterText, 0, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
                    .addGroup(layout.createSequentialGroup()
                        .addComponent(jLabel1)
                        .addGap(0, 168, Short.MAX_VALUE))
                    .addGroup(layout.createSequentialGroup()
                        .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                            .addComponent(button1, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, 158, Short.MAX_VALUE)
                            .addComponent(StartButton, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 526, Short.MAX_VALUE)
                        .addComponent(proceedButton, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addContainerGap())
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jLabel1)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(jLabel4)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(button1, javax.swing.GroupLayout.PREFERRED_SIZE, 40, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(specSheetPath, javax.swing.GroupLayout.PREFERRED_SIZE, 40, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(button2, javax.swing.GroupLayout.PREFERRED_SIZE, 40, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(machineDataPath, javax.swing.GroupLayout.PREFERRED_SIZE, 40, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(30, 30, 30)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jLabel2)
                    .addComponent(jLabel3)
                    .addComponent(modelNumberText, javax.swing.GroupLayout.PREFERRED_SIZE, 37, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(screwDiameterText, javax.swing.GroupLayout.PREFERRED_SIZE, 37, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(46, 46, 46)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(proceedButton, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(StartButton, javax.swing.GroupLayout.PREFERRED_SIZE, 39, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addGap(18, 18, 18)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 283, Short.MAX_VALUE)
                .addContainerGap())
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void StartButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_StartButtonActionPerformed

        modelNumber = (String)modelNumberText.getSelectedItem();
        screwDiameter = (String)screwDiameterText.getSelectedItem();

        Workbook w;
        try{
            w = Workbook.getWorkbook(specSheet);
            Sheet sheet = w.getSheet(2);
            //Checking for Model Number
            for (int i = 0; i < sheet.getColumns(); i++){
                Cell cell = sheet.getCell(i, 2);
                CellType type = cell.getType();
                if (modelNumber.equals(cell.getContents())){
                    modelColomnNumber = i;
                }
            }

            //Checking for screw diameter
            if (modelColomnNumber != -1){
                for (int i = 0; i < 4; i++){
                    Cell cell = sheet.getCell(modelColomnNumber + i, 8);
                    if (cell.getContents().equals(screwDiameter)){
                        colomnNumber = modelColomnNumber + i;
                    }
                }
            }

            //Printing out values
            if (colomnNumber != -1 && modelColomnNumber != -1){
                for (int i = 462; i < sheet.getRows(); i++){

                    Cell varCell = sheet.getCell(1, i);
                    Cell rawCell = sheet.getCell(modelColomnNumber, i);
                    Cell cell = sheet.getCell(colomnNumber, i);
                    CellType type = cell.getType();

                    if (!cell.getContents().equals("") && varCell.getContents().matches(".*[a-zA-Z]+.*")){
                        if (cell.getContents().equals("YES") || cell.getContents().equals("NO")){
                            if (cell.getContents().equals("YES"))
                            specSheetValues.add("true");
                            if (cell.getContents().equals("NO"))
                            specSheetValues.add("false");    
                        }
                        else
                        specSheetValues.add(rawCell.getContents().replace(",", ""));
                        specSheetVariableNames.add(varCell.getContents());
                    }
                    else if (!rawCell.getContents().equals("") && varCell.getContents().matches(".*[a-zA-Z]+.*")){
                        specSheetValues.add(rawCell.getContents().replace(",", ""));
                        specSheetVariableNames.add(varCell.getContents());
                    }
                }
            }

            if (colomnNumber == -1){
                area.append("**********************************\n");
                area.append("Invalid Model Number or Screw Size\n");
                area.append("**********************************\n");
            }

        }catch(BiffException e){
            e.printStackTrace();
        } catch (IOException ex) {
            Logger.getLogger(GUI.class.getName()).log(Level.SEVERE, null, ex);
        }

        //XML portion

        try {

            DocumentBuilderFactory dbFactory = DocumentBuilderFactory.newInstance();
            DocumentBuilder dBuilder = dbFactory.newDocumentBuilder();
            Document doc = dBuilder.parse(fXmlFile);

            //optional, but recommended
            //read this - http://stackoverflow.com/questions/13786607/normalization-in-dom-parsing-with-java-how-does-it-work
            doc.getDocumentElement().normalize();

            area.append("Root element :" + doc.getDocumentElement().getNodeName() + "\n");

            NodeList nList = doc.getElementsByTagName("Variable");

            area.append("----------------------------\n");

            for (int temp = 0; temp < nList.getLength(); temp++) {

                Node nNode = nList.item(temp);

                Element eElement = (Element) nNode;
                
                xmlVariableNames.add(eElement.getAttribute("Name"));
                xmlValues.add(eElement.getElementsByTagName("Value").item(0).getTextContent());


            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        
        for (int i = 0; i < xmlValues.size(); i++){
            if (xmlVariableNames.get(i).contains("Injection1.sv_ScrewLintab.LintabPoints.Point")){
                xmlValues.set(i, xmlValues.get(i).replace(",", ""));
                float Temp = Float.parseFloat(xmlValues.get(i));
                Temp /= 10;
                xmlValues.set((i), String.valueOf(Temp));
            }
        }
        
        for (int i = 0; i < specSheetValues.size(); i++){
            for (int j = 0; j < xmlValues.size(); j++){
                if (specSheetVariableNames.get(i).equals(xmlVariableNames.get(j))){
                    if (specSheetValues.get(i).matches(".*[a-zA-Z]+.*")){
                        if (specSheetValues.get(i).equals(xmlValues.get(j))){
                            area.append("Correct\n");
                            area.append("----------------------------------------------------\n");
                        }
                        else{
                            area.append(specSheetVariableNames.get(i) + "\n");
                            area.append("Spec Sheet:   " + specSheetValues.get(i) + "\n");
                            area.append("Machine Data: " + xmlValues.get(j) + "\n");
                            area.append("----------------------------------------------------\n");
                            mistakes.add(specSheetVariableNames.get(i));
                        }
                    }
                    else{
                        if (Double.parseDouble(specSheetValues.get(i).replaceAll(",", "")) == Double.parseDouble(xmlValues.get(j).replaceAll(",", ""))){
                            area.append("Correct\n");
                            area.append("----------------------------------------------------\n");
                        }
                        else{
                            area.append(specSheetVariableNames.get(i) + "\n");
                            area.append("Spec Sheet:   " + specSheetValues.get(i) + "\n");
                            area.append("Machine Data: " + xmlValues.get(j) + "\n");
                            area.append("----------------------------------------------------\n");
                            mistakes.add(specSheetVariableNames.get(i));
                        }
                    }
                    //												System.out.println(specSheetValues.get(i) + " Compared to " + xmlValues.get(j));
                    //												System.out.println(specSheetVariableNames.get(i));
                    //												System.out.println();
                }
            }
        }
    }//GEN-LAST:event_StartButtonActionPerformed

    private void button2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_button2ActionPerformed
        JFileChooser chooser = new JFileChooser();
        chooser.setCurrentDirectory(new File("."));

        System.out.println("Select the machine data");
        int r = chooser.showOpenDialog(new JFrame());

        fXmlFile = chooser.getSelectedFile().getAbsoluteFile();
        machineDataPath.setText(chooser.getSelectedFile().getAbsolutePath());
    }//GEN-LAST:event_button2ActionPerformed

    private void button1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_button1ActionPerformed
        JFileChooser chooser = new JFileChooser();
        chooser.setCurrentDirectory(new File("."));

        int r = chooser.showOpenDialog(new JFrame());
        //path = chooser.getSelectedFile().getName();

        specSheet = chooser.getSelectedFile().getAbsoluteFile();
        specSheetPath.setText(chooser.getSelectedFile().getAbsolutePath());

    }//GEN-LAST:event_button1ActionPerformed

    private void proceedButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_proceedButtonActionPerformed
        ChangeXML changeXml = new ChangeXML(specSheetVariableNames, specSheetValues, xmlVariableNames, xmlValues, mistakes, fXmlFile);
        changeXml.setVisible(true);
        System.out.println("Worked");
        //WindowEvent wev = new WindowEvent(this, WindowEvent.WINDOW_CLOSING);
        //Toolkit.getDefaultToolkit().getSystemEventQueue().postEvent(wev);
    }//GEN-LAST:event_proceedButtonActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(GUI.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(GUI.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(GUI.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(GUI.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new GUI().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private java.awt.Button StartButton;
    private javax.swing.JTextArea area;
    private java.awt.Button button1;
    private java.awt.Button button2;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTextField machineDataPath;
    private javax.swing.JComboBox modelNumberText;
    private java.awt.Button proceedButton;
    private javax.swing.JComboBox screwDiameterText;
    private javax.swing.JTextField specSheetPath;
    // End of variables declaration//GEN-END:variables
}
