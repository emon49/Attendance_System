/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package javaapplication2;

import java.awt.Toolkit;
import java.awt.print.PrinterException;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.MessageFormat;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;
import static org.apache.poi.hemf.record.emf.HemfRecordType.header;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author HP
 */
public class overallReport extends javax.swing.JFrame {

    /**
     * Creates new form overallReport
     */
    String course,selectedStatus;
    int local;
    DefaultTableModel model; 
    public overallReport() {
        initComponents();
        setIcon();
    }
    
    public overallReport(String cou,int loc) {
        initComponents();
        course=cou;
        local=loc;
        setIcon();
    }
    
    void showOverall(){
        if(course.equals("JAVA")){
             ArrayList<String> nameList=new ArrayList<String>();
             ArrayList<String> rollList=new ArrayList<String>();
             ArrayList<Float> persentageList=new ArrayList<Float>();
             
              
            Workbook workbook;
            try {
                workbook = new XSSFWorkbook(new FileInputStream("java_Attendance.xlsx"));
                 Sheet sheet=workbook.getSheetAt(0);
                int rowCount= sheet.getPhysicalNumberOfRows();
                int columCount=sheet.getRow(0).getLastCellNum();
                //XSSFRow row1=(XSSFRow) sheet.getRow(0);
                for(int i=1;i<rowCount;i++)
                    {
                        int present=0;
                        XSSFRow row=(XSSFRow) sheet.getRow(i);
                        String roll=row.getCell(0).toString();
                        String name=row.getCell(1).toString();
                        
                        for(int j=2;j<columCount;j++)
                        {
                            
                            String state=row.getCell(j).toString();
                            if(state.equals("Pr"))
                            {
                                present++;
                            }
                        }
                        int total=columCount-2;
                        float persentage=(float) ((present/ (1.0*total) )*100);
                        rollList.add(roll);
                        nameList.add(name);
                        persentageList.add(persentage);     
                    }
                
                
            } catch (FileNotFoundException ex) {
                Logger.getLogger(overallReport.class.getName()).log(Level.SEVERE, null, ex);
            } catch (IOException ex) {
                Logger.getLogger(overallReport.class.getName()).log(Level.SEVERE, null, ex);
            }
               
            if(selectedStatus.equals("Non Collegiate"))
            {
                 model=(DefaultTableModel) jTable1.getModel(); 
                 model.setRowCount(0);
                 for(int i=0;i<persentageList.size();i++)
                {
                    if(persentageList.get(i)>=60.0 && persentageList.get(i)<=75.0)
                    {
                         model.insertRow(model.getRowCount(),new Object[]{nameList.get(i),rollList.get(i),persentageList.get(i)});
                    }
                        
                }
            }
             if(selectedStatus.equals("Dis Collegiate"))
            {
                 model=(DefaultTableModel) jTable1.getModel(); 
                 model.setRowCount(0);
                 for(int i=0;i<persentageList.size();i++)
                {
                    if(persentageList.get(i)<60.0)
                    {
                         model.insertRow(model.getRowCount(),new Object[]{nameList.get(i),rollList.get(i),persentageList.get(i)});
                    }
                        
                }
            }
            if(selectedStatus.equals("Collegiate"))
            {
                 model=(DefaultTableModel) jTable1.getModel();
                 model.setRowCount(0);
                 for(int i=0;i<persentageList.size();i++)
                {
                    if(persentageList.get(i)>75.0)
                    {
                         model.insertRow(model.getRowCount(),new Object[]{nameList.get(i),rollList.get(i),persentageList.get(i)});
                    }   
                }
            }
//            Collegiate
//            Non Collegiate
//            Dis Collegiate
        }
    }

    
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        jScrollPane1 = new javax.swing.JScrollPane();
        jTable1 = new javax.swing.JTable();
        jLabel1 = new javax.swing.JLabel();
        jComboBox1 = new javax.swing.JComboBox<>();
        jButton1 = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setMaximumSize(new java.awt.Dimension(930, 750));
        setMinimumSize(new java.awt.Dimension(930, 750));
        setPreferredSize(new java.awt.Dimension(930, 750));
        getContentPane().setLayout(null);

        jPanel1.setBackground(new java.awt.Color(44, 44, 84));
        jPanel1.setLayout(null);

        jTable1.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "Name", "Roll", "Status"
            }
        ));
        jScrollPane1.setViewportView(jTable1);

        jPanel1.add(jScrollPane1);
        jScrollPane1.setBounds(80, 100, 760, 520);

        jLabel1.setFont(new java.awt.Font("Segoe UI", 1, 24)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(255, 255, 255));
        jLabel1.setText("Catagory:");
        jPanel1.add(jLabel1);
        jLabel1.setBounds(200, 20, 130, 50);

        jComboBox1.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jComboBox1.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Collegiate", "Non Collegiate", "Dis Collegiate" }));
        jComboBox1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jComboBox1ActionPerformed(evt);
            }
        });
        jPanel1.add(jComboBox1);
        jComboBox1.setBounds(330, 30, 170, 40);

        jButton1.setText("Export As PDF");
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });
        jPanel1.add(jButton1);
        jButton1.setBounds(350, 655, 240, 30);

        getContentPane().add(jPanel1);
        jPanel1.setBounds(0, 0, 930, 750);

        pack();
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void jComboBox1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jComboBox1ActionPerformed
       selectedStatus= jComboBox1.getSelectedItem().toString();
       showOverall();
    }//GEN-LAST:event_jComboBox1ActionPerformed

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        try{
      
       
       if(selectedStatus.equals("Non Collegiate"))
       {
           MessageFormat header=new MessageFormat("List of Non Collegiate Students");
           MessageFormat footer=new MessageFormat("Page{0,number,integer}");
           jTable1.print(JTable.PrintMode.NORMAL,header,footer);
       }
       else if(selectedStatus.equals("Dis Collegiate"))
       {
           MessageFormat header=new MessageFormat("List of Dis Collegiate Students");
           MessageFormat footer=new MessageFormat("Page{0,number,integer}");
           jTable1.print(JTable.PrintMode.NORMAL,header,footer);
       }
       else if(selectedStatus.equals("Collegiate"))
       {
           MessageFormat header=new MessageFormat("List of Collegiate Students");
           MessageFormat footer=new MessageFormat("Page{0,number,integer}");
           jTable1.print(JTable.PrintMode.NORMAL,header,footer);
       }
     }
        catch(Exception e){
            
        }
    }//GEN-LAST:event_jButton1ActionPerformed

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
            java.util.logging.Logger.getLogger(overallReport.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(overallReport.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(overallReport.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(overallReport.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new overallReport().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton1;
    private javax.swing.JComboBox<String> jComboBox1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTable jTable1;
    // End of variables declaration//GEN-END:variables
    private void setIcon() {
        setIconImage(Toolkit.getDefaultToolkit().getImage(getClass().getResource("mist_logo.png")));
    }

}
