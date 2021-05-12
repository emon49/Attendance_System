/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package javaapplication2;

import java.awt.Toolkit;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class studentReport extends javax.swing.JFrame {

    String cours;
    int local;
    int present=0,absent=0,total=0,percentage=0;
    public studentReport() {
        initComponents();
        subPanel.setVisible(false);
        setIcon();
    }
    
    public studentReport(String cou,int row)
    {
        setIcon();
         initComponents();
          subPanel.setVisible(false);
          cours=cou;
          local=row;   
    }
    
    
    
    
    

     
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        mainPanel = new javax.swing.JPanel();
        subPanel = new javax.swing.JPanel();
        reportLabel = new javax.swing.JLabel();
        nameLabel = new javax.swing.JLabel();
        jLabel1 = new javax.swing.JLabel();
        totalClassLabel = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        presentClasses = new javax.swing.JLabel();
        jLabel5 = new javax.swing.JLabel();
        absentClass = new javax.swing.JLabel();
        jLabel7 = new javax.swing.JLabel();
        stateLabel = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        rollField = new javax.swing.JTextField();
        goBack = new javax.swing.JLabel();
        reportButton = new javax.swing.JLabel();
        jLabel6 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        mainPanel.setBackground(new java.awt.Color(44, 44, 84));
        mainPanel.setLayout(null);

        subPanel.setBackground(new java.awt.Color(44, 44, 84));
        subPanel.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 204, 204), 3));
        subPanel.setLayout(null);

        reportLabel.setFont(new java.awt.Font("Segoe UI", 1, 24)); // NOI18N
        reportLabel.setForeground(new java.awt.Color(255, 255, 255));
        reportLabel.setText("Report of  ");
        subPanel.add(reportLabel);
        reportLabel.setBounds(250, 20, 150, 50);

        nameLabel.setFont(new java.awt.Font("Segoe UI", 1, 24)); // NOI18N
        nameLabel.setForeground(new java.awt.Color(255, 255, 255));
        subPanel.add(nameLabel);
        nameLabel.setBounds(370, 20, 370, 50);

        jLabel1.setFont(new java.awt.Font("Segoe UI", 1, 24)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(255, 255, 255));
        jLabel1.setText("Total Classes:");
        subPanel.add(jLabel1);
        jLabel1.setBounds(70, 90, 170, 30);

        totalClassLabel.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        totalClassLabel.setForeground(new java.awt.Color(255, 255, 255));
        subPanel.add(totalClassLabel);
        totalClassLabel.setBounds(320, 90, 90, 30);

        jLabel3.setFont(new java.awt.Font("Segoe UI", 1, 24)); // NOI18N
        jLabel3.setForeground(new java.awt.Color(255, 255, 255));
        jLabel3.setText("Present :");
        subPanel.add(jLabel3);
        jLabel3.setBounds(70, 150, 160, 30);

        presentClasses.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        presentClasses.setForeground(new java.awt.Color(255, 255, 255));
        subPanel.add(presentClasses);
        presentClasses.setBounds(320, 150, 80, 30);

        jLabel5.setFont(new java.awt.Font("Segoe UI", 1, 24)); // NOI18N
        jLabel5.setForeground(new java.awt.Color(255, 255, 255));
        jLabel5.setText("Absent :");
        subPanel.add(jLabel5);
        jLabel5.setBounds(70, 210, 110, 40);

        absentClass.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        absentClass.setForeground(new java.awt.Color(255, 255, 255));
        subPanel.add(absentClass);
        absentClass.setBounds(320, 220, 100, 30);

        jLabel7.setFont(new java.awt.Font("Segoe UI", 1, 24)); // NOI18N
        jLabel7.setForeground(new java.awt.Color(255, 255, 255));
        jLabel7.setText("Attendance State: ");
        subPanel.add(jLabel7);
        jLabel7.setBounds(70, 280, 240, 30);

        stateLabel.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        stateLabel.setForeground(new java.awt.Color(255, 255, 255));
        subPanel.add(stateLabel);
        stateLabel.setBounds(310, 280, 200, 40);

        mainPanel.add(subPanel);
        subPanel.setBounds(40, 240, 890, 350);

        jLabel2.setFont(new java.awt.Font("Segoe UI", 1, 24)); // NOI18N
        jLabel2.setForeground(new java.awt.Color(255, 255, 255));
        jLabel2.setText("Enter Roll:");
        mainPanel.add(jLabel2);
        jLabel2.setBounds(90, 90, 130, 50);

        rollField.setFont(new java.awt.Font("Segoe UI", 0, 24)); // NOI18N
        rollField.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                rollFieldActionPerformed(evt);
            }
        });
        mainPanel.add(rollField);
        rollField.setBounds(300, 90, 180, 40);

        goBack.setIcon(new javax.swing.ImageIcon(getClass().getResource("/javaapplication2/gobackStudentreport.png"))); // NOI18N
        goBack.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        goBack.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                goBackMouseClicked(evt);
            }
        });
        mainPanel.add(goBack);
        goBack.setBounds(821, 20, 100, 35);

        reportButton.setIcon(new javax.swing.ImageIcon(getClass().getResource("/javaapplication2/report.png"))); // NOI18N
        reportButton.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                reportButtonMouseClicked(evt);
            }
        });
        mainPanel.add(reportButton);
        reportButton.setBounds(530, 80, 120, 60);

        jLabel6.setIcon(new javax.swing.ImageIcon(getClass().getResource("/javaapplication2/overall.png"))); // NOI18N
        jLabel6.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        jLabel6.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jLabel6MouseClicked(evt);
            }
        });
        mainPanel.add(jLabel6);
        jLabel6.setBounds(700, 90, 190, 50);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addComponent(mainPanel, javax.swing.GroupLayout.PREFERRED_SIZE, 958, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(0, 0, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addComponent(mainPanel, javax.swing.GroupLayout.PREFERRED_SIZE, 623, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(0, 0, Short.MAX_VALUE))
        );

        pack();
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents
       
    private void reportButtonMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_reportButtonMouseClicked
       
        
        String roll=rollField.getText();
        //Extrtacting year of roll
        String first ="";
       for(int i=0;i<4;i++)
       {
          // char ch=roll[i];
           
           first=first+roll.charAt(i);
       }
        //System.out.println(first);
        
        
        
        
         int r = 0;
        if(cours.equals("JAVA"))
        {
            
            try{
                if(!(first.equals("2019")))
                {
                    throw new rollNotFoundException("Student of roll "+roll+" haven't take this course");
                }
                
                else{
                    Workbook workbook;
            try {
                workbook = new XSSFWorkbook(new FileInputStream("java_Attendance.xlsx"));
                Sheet sheet=workbook.getSheetAt(0);
                 int rowCount= sheet.getPhysicalNumberOfRows();
                 for(int i=1;i<rowCount;i++)
                 {
                     XSSFRow row=(XSSFRow) sheet.getRow(i);
                    String studentRoll=row.getCell(0).toString();
                   
                     if(roll.equals(studentRoll))
                     {
                         r=i;
                         break;
                     }
                 }  ///break of for loop
                 
                 //Cell name=sh.getRow(r).getCell(0);
                //  String nm=name.getRichStringCellValue().toString();
                
                XSSFRow row=(XSSFRow) sheet.getRow(r);
                //Taking student name
                String studentName=row.getCell(1).toString();
                int columCount=sheet.getRow(r).getLastCellNum();
                
                //loop for counting
                
                for(int k=2;k<columCount;k++)
                {
                     String att=row.getCell(k).toString();
                    if(att.equals("Pr"))
                    {
                        present++;
                    }
                    else
                    {
                        absent++;
                    }
                    
                    
                }
                total=present+absent;
                
                String to=Integer.toString(total);
                String pre=Integer.toString(present);
                String Abs=Integer.toString(absent);
                
                nameLabel.setText(studentName);
                totalClassLabel.setText(to);
                presentClasses.setText(pre);
                absentClass.setText(Abs);
                
                
                percentage=(present*100)/total;
                
                if(percentage>75)
                {
                    stateLabel.setText("Collegiate");
                }
                else if(percentage>=60)
                {
                    stateLabel.setText("Non-Collegiate");
                }
                
                else
                {
                    stateLabel.setText("Dis-Collegiate");
                }
                
                subPanel.setVisible(true);
                //System.out.println("Present: "+present+" "+"Absent: "+absent);
              
            } catch (FileNotFoundException ex) {
                Logger.getLogger(studentReport.class.getName()).log(Level.SEVERE, null, ex);
            } catch (IOException ex) {
                Logger.getLogger(studentReport.class.getName()).log(Level.SEVERE, null, ex);
            }
           
                }
            }catch(rollNotFoundException ex)
            {
                JOptionPane.showMessageDialog(null,ex,"Not Found",JOptionPane.WARNING_MESSAGE);
            }
           
        }
        
        
        
        
           //subPanel.setVisible(true);
         
    }//GEN-LAST:event_reportButtonMouseClicked

    
    
    
  
    private void rollFieldActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_rollFieldActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_rollFieldActionPerformed

    private void goBackMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_goBackMouseClicked
        this.setVisible(false);
        ChoosingPage csp;
        try {
            csp = new ChoosingPage(local);
            csp.setVisible(true);
        } catch (IOException ex) {
            Logger.getLogger(studentReport.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }//GEN-LAST:event_goBackMouseClicked

    private void jLabel6MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel6MouseClicked
       this.setVisible(false);
       overallReport or=new overallReport(cours,local);
       or.setVisible(true);
    }//GEN-LAST:event_jLabel6MouseClicked

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
            java.util.logging.Logger.getLogger(studentReport.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(studentReport.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(studentReport.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(studentReport.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new studentReport().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JLabel absentClass;
    private javax.swing.JLabel goBack;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JPanel mainPanel;
    private javax.swing.JLabel nameLabel;
    private javax.swing.JLabel presentClasses;
    private javax.swing.JLabel reportButton;
    private javax.swing.JLabel reportLabel;
    private javax.swing.JTextField rollField;
    private javax.swing.JLabel stateLabel;
    private javax.swing.JPanel subPanel;
    private javax.swing.JLabel totalClassLabel;
    // End of variables declaration//GEN-END:variables
private void setIcon() {
        setIconImage(Toolkit.getDefaultToolkit().getImage(getClass().getResource("mist_logo.png")));
    }

}
