
package javaapplication2;

import com.toedter.calendar.JDateChooser;
import java.awt.Toolkit;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ChoosingPage extends javax.swing.JFrame {

    SimpleDateFormat format1;
    Workbook wb;
    Sheet sh;
    int localrow;
    String selectedCourse;
    String intdate;
    public ChoosingPage() {
        initComponents();
        setIcon();
    }
    
    // Parameterized constructor
    public ChoosingPage(int row) throws IOException {
        setIcon();
        ///System.out.println(row);
        
        initComponents();
         localrow=row;

        showLoginName(localrow);
       /// System.out.println("local row:"+localrow);
        
    }

    
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        ChoosingPanel = new javax.swing.JPanel();
        teacherNameLabel = new javax.swing.JLabel();
        DeptChooseLabel = new javax.swing.JLabel();
        CoarseChooseLabel = new javax.swing.JLabel();
        TimeChooseLabel = new javax.swing.JLabel();
        deptComboBox = new javax.swing.JComboBox<>();
        coarseComboBox = new javax.swing.JComboBox<>();
        jDateChooser = new com.toedter.calendar.JDateChooser();
        attendanceMouse = new javax.swing.JLabel();
        studentReport = new javax.swing.JLabel();
        logoutButton = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        ChoosingPanel.setBackground(new java.awt.Color(118, 194, 175));
        ChoosingPanel.setPreferredSize(new java.awt.Dimension(952, 784));
        ChoosingPanel.setLayout(null);

        teacherNameLabel.setFont(new java.awt.Font("Bodoni MT", 3, 36)); // NOI18N
        teacherNameLabel.setForeground(new java.awt.Color(255, 255, 255));
        ChoosingPanel.add(teacherNameLabel);
        teacherNameLabel.setBounds(380, 110, 369, 70);

        DeptChooseLabel.setFont(new java.awt.Font("Calibri", 1, 24)); // NOI18N
        DeptChooseLabel.setForeground(new java.awt.Color(255, 255, 255));
        DeptChooseLabel.setText("DEPT");
        ChoosingPanel.add(DeptChooseLabel);
        DeptChooseLabel.setBounds(250, 180, 80, 50);

        CoarseChooseLabel.setFont(new java.awt.Font("Calibri", 1, 24)); // NOI18N
        CoarseChooseLabel.setForeground(new java.awt.Color(255, 255, 255));
        CoarseChooseLabel.setText("Coarse");
        ChoosingPanel.add(CoarseChooseLabel);
        CoarseChooseLabel.setBounds(250, 290, 80, 50);

        TimeChooseLabel.setFont(new java.awt.Font("Calibri", 1, 24)); // NOI18N
        TimeChooseLabel.setForeground(new java.awt.Color(255, 255, 255));
        TimeChooseLabel.setText("Date");
        ChoosingPanel.add(TimeChooseLabel);
        TimeChooseLabel.setBounds(250, 390, 60, 60);

        deptComboBox.setFont(new java.awt.Font("Calibri", 1, 18)); // NOI18N
        deptComboBox.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "CSE", "EECE", "NSE", "BME", "PME", "EWCE", "ARCHI" }));
        deptComboBox.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                deptComboBoxActionPerformed(evt);
            }
        });
        ChoosingPanel.add(deptComboBox);
        deptComboBox.setBounds(400, 190, 210, 30);

        coarseComboBox.setFont(new java.awt.Font("Calibri", 1, 18)); // NOI18N
        coarseComboBox.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                coarseComboBoxActionPerformed(evt);
            }
        });
        ChoosingPanel.add(coarseComboBox);
        coarseComboBox.setBounds(400, 300, 210, 30);
        ChoosingPanel.add(jDateChooser);
        jDateChooser.setBounds(400, 400, 210, 30);

        attendanceMouse.setIcon(new javax.swing.ImageIcon(getClass().getResource("/javaapplication2/ClassAttendance.png"))); // NOI18N
        attendanceMouse.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                attendanceMouseMouseClicked(evt);
            }
        });
        ChoosingPanel.add(attendanceMouse);
        attendanceMouse.setBounds(700, 640, 220, 110);

        studentReport.setIcon(new javax.swing.ImageIcon(getClass().getResource("/javaapplication2/StudentReport.png"))); // NOI18N
        studentReport.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                studentReportMouseClicked(evt);
            }
        });
        ChoosingPanel.add(studentReport);
        studentReport.setBounds(20, 660, 240, 50);

        logoutButton.setIcon(new javax.swing.ImageIcon(getClass().getResource("/javaapplication2/logout_choosingpage.png"))); // NOI18N
        logoutButton.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        logoutButton.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                logoutButtonMouseClicked(evt);
            }
        });
        ChoosingPanel.add(logoutButton);
        logoutButton.setBounds(830, 10, 90, 40);

        jLabel2.setFont(new java.awt.Font("Tahoma", 0, 18)); // NOI18N
        jLabel2.setIcon(new javax.swing.ImageIcon(getClass().getResource("/javaapplication2/computer.png"))); // NOI18N
        ChoosingPanel.add(jLabel2);
        jLabel2.setBounds(-30, 0, 990, 780);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(ChoosingPanel, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(ChoosingPanel, javax.swing.GroupLayout.DEFAULT_SIZE, 782, Short.MAX_VALUE)
        );

        pack();
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void deptComboBoxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_deptComboBoxActionPerformed
  
    }//GEN-LAST:event_deptComboBoxActionPerformed

    private void coarseComboBoxActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_coarseComboBoxActionPerformed
        selectedCourse=coarseComboBox.getSelectedItem().toString();
    }//GEN-LAST:event_coarseComboBoxActionPerformed

    private void attendanceMouseMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_attendanceMouseMouseClicked
        Date date=null;
        date = jDateChooser.getDate();
        
        format1=new SimpleDateFormat("MMM dd,yyyy");
        
        try{
            //convert date to string
            intdate=format1.format(date);
            //System.out.println(intdate);              
        }
        catch(Exception e)
        {
            System.out.println("Exception Called");
        }              
            this.setVisible(false);
            attendanceSheet as;
            as=new attendanceSheet(selectedCourse,intdate);
            as.setVisible(true);
    }//GEN-LAST:event_attendanceMouseMouseClicked

    private void studentReportMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_studentReportMouseClicked
        studentReport st=new studentReport(selectedCourse,localrow);
        this.setVisible(false);
        st.setVisible(true);
    }//GEN-LAST:event_studentReportMouseClicked

    private void logoutButtonMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_logoutButtonMouseClicked
       this.setVisible(false);
       LogIn ln=new LogIn();
       ln.setVisible(true);
    }//GEN-LAST:event_logoutButtonMouseClicked

    private void showLoginName(int lrow) throws FileNotFoundException, IOException
    {
        /* To show Teacher's name +add course item in Combo Box  */
        wb=new XSSFWorkbook(new FileInputStream("teacherInfo.xlsx"));
        sh=wb.getSheetAt(0);
        Cell name=sh.getRow(lrow).getCell(0);
        String nm=name.getRichStringCellValue().toString();
        
        
        teacherNameLabel.setText(nm);
        
        
        int columCount=sh.getRow(lrow).getLastCellNum();
           ///System.out.println(columCount);
           int j=0;
           for(int i=3;i<columCount;i++)
           {
               Cell nam=sh.getRow(lrow).getCell(i);
               String s;
               
             
               s = nam.getRichStringCellValue().toString();
              // System.out.println(s);

                coarseComboBox.addItem(s);
              
           }
           
           
           
           ///For date
       
        
        Date todaysdate =new Date();
        jDateChooser.setDate(todaysdate);  

        
    }
    
  

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
            java.util.logging.Logger.getLogger(ChoosingPage.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(ChoosingPage.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(ChoosingPage.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(ChoosingPage.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new ChoosingPage().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JPanel ChoosingPanel;
    private javax.swing.JLabel CoarseChooseLabel;
    private javax.swing.JLabel DeptChooseLabel;
    private javax.swing.JLabel TimeChooseLabel;
    private javax.swing.JLabel attendanceMouse;
    private javax.swing.JComboBox<String> coarseComboBox;
    private javax.swing.JComboBox<String> deptComboBox;
    private com.toedter.calendar.JDateChooser jDateChooser;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel logoutButton;
    private javax.swing.JLabel studentReport;
    private javax.swing.JLabel teacherNameLabel;
    // End of variables declaration//GEN-END:variables

     private void setIcon() {
        setIconImage(Toolkit.getDefaultToolkit().getImage(getClass().getResource("mist_logo.png")));
    }
}
