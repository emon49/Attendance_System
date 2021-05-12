/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package javaapplication2;

import java.awt.Toolkit;
import java.io.FileInputStream;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.text.html.HTMLDocument;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.xssf.usermodel.XSSFRow;
public class LogIn extends javax.swing.JFrame {

   private static LogIn l;
    public LogIn() {
        initComponents();
        setIcon();
    }

   
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jLabel4 = new javax.swing.JLabel();
        rightPanel = new javax.swing.JPanel();
        leftPanel = new javax.swing.JPanel();
        userIconLeft = new javax.swing.JLabel();
        logInButton = new javax.swing.JLabel();
        userIconRight = new javax.swing.JLabel();
        userPassword = new javax.swing.JLabel();
        user_name = new javax.swing.JLabel();
        userNameTextField = new javax.swing.JTextField();
        password = new javax.swing.JLabel();
        passwordTextField = new javax.swing.JPasswordField();

        jLabel4.setText("jLabel4");

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setPreferredSize(new java.awt.Dimension(768, 777));
        setResizable(false);

        rightPanel.setBackground(new java.awt.Color(255, 255, 255));
        rightPanel.setPreferredSize(new java.awt.Dimension(691, 777));
        rightPanel.setLayout(null);

        leftPanel.setBackground(new java.awt.Color(24, 44, 97));

        userIconLeft.setIcon(new javax.swing.ImageIcon(getClass().getResource("/javaapplication2/circle-cropped.png"))); // NOI18N

        javax.swing.GroupLayout leftPanelLayout = new javax.swing.GroupLayout(leftPanel);
        leftPanel.setLayout(leftPanelLayout);
        leftPanelLayout.setHorizontalGroup(
            leftPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(leftPanelLayout.createSequentialGroup()
                .addGap(44, 44, 44)
                .addComponent(userIconLeft)
                .addContainerGap(80, Short.MAX_VALUE))
        );
        leftPanelLayout.setVerticalGroup(
            leftPanelLayout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(leftPanelLayout.createSequentialGroup()
                .addGap(167, 167, 167)
                .addComponent(userIconLeft)
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );

        rightPanel.add(leftPanel);
        leftPanel.setBounds(0, 0, 344, 777);
        leftPanel.getAccessibleContext().setAccessibleName("");

        logInButton.setIcon(new javax.swing.ImageIcon(getClass().getResource("/javaapplication2/button1.png"))); // NOI18N
        logInButton.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        logInButton.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                logInButtonMouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                logInButtonMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                logInButtonMouseExited(evt);
            }
        });
        rightPanel.add(logInButton);
        logInButton.setBounds(503, 392, 154, 87);
        logInButton.getAccessibleContext().setAccessibleName("Button2");

        userIconRight.setIcon(new javax.swing.ImageIcon(getClass().getResource("/javaapplication2/user_name2.png"))); // NOI18N
        rightPanel.add(userIconRight);
        userIconRight.setBounds(393, 151, 64, 78);

        userPassword.setIcon(new javax.swing.ImageIcon(getClass().getResource("/javaapplication2/pass2.png"))); // NOI18N
        rightPanel.add(userPassword);
        userPassword.setBounds(414, 339, 24, 24);

        user_name.setFont(new java.awt.Font("Segoe UI", 1, 36)); // NOI18N
        user_name.setText("User Id");
        rightPanel.add(user_name);
        user_name.setBounds(510, 80, 134, 50);

        userNameTextField.setFont(new java.awt.Font("Segoe UI", 1, 24)); // NOI18N
        userNameTextField.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                userNameTextFieldActionPerformed(evt);
            }
        });
        rightPanel.add(userNameTextField);
        userNameTextField.setBounds(475, 151, 208, 67);

        password.setFont(new java.awt.Font("Segoe UI", 1, 36)); // NOI18N
        password.setText("Password");
        rightPanel.add(password);
        password.setBounds(490, 230, 190, 50);

        passwordTextField.setFont(new java.awt.Font("Segoe UI", 1, 24)); // NOI18N
        rightPanel.add(passwordTextField);
        passwordTextField.setBounds(475, 303, 208, 58);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(rightPanel, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.PREFERRED_SIZE, 768, javax.swing.GroupLayout.PREFERRED_SIZE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(rightPanel, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );

        getAccessibleContext().setAccessibleName("LogInJframe");

        pack();
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void logInButtonMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_logInButtonMouseEntered
        logInButton.setIcon(new javax.swing.ImageIcon(getClass().getResource("./button2.png")));
    }//GEN-LAST:event_logInButtonMouseEntered

    private void logInButtonMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_logInButtonMouseExited
        logInButton.setIcon(new javax.swing.ImageIcon(getClass().getResource("./button1.png")));
    }//GEN-LAST:event_logInButtonMouseExited

    private void userNameTextFieldActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_userNameTextFieldActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_userNameTextFieldActionPerformed

    private void logInButtonMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_logInButtonMouseClicked
       String s1=userNameTextField.getText();
       String s2=passwordTextField.getText();
       
       
                           
        try {
            Workbook workbook=new XSSFWorkbook(new FileInputStream("teacherInfo.xlsx"));
            Sheet sheet=workbook.getSheetAt(0);
            int rowCount= sheet.getPhysicalNumberOfRows();
            ////System.out.println(rowCount);
            int columCount=sheet.getRow(0).getLastCellNum();
            int i;
            for( i=0;i<rowCount;i++)
                {
                    XSSFRow row=(XSSFRow) sheet.getRow(i);
                    String userID=row.getCell(1).toString();
                    String pass=row.getCell(2).toString();
                   
                    
                    if(s1.equals(userID) && s2.equals(pass))
                    {
                           ChoosingPage ch;
                           ch = new ChoosingPage(i);

                           this.setVisible(false);
                           ch.setVisible(true);
                          
                          ///JOptionPane.showMessageDialog(null,"Successful","LogIn Info",JOptionPane.INFORMATION_MESSAGE);

                      

                          break;
                        
                          
                    }
                   
                }
            
            if(i==rowCount)
            {
                JOptionPane.showMessageDialog(null,"Invalid User","LogIn Info",JOptionPane.WARNING_MESSAGE);
            }
            
            
            

        } catch (FileNotFoundException ex) {
            Logger.getLogger(LogIn.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(LogIn.class.getName()).log(Level.SEVERE, null, ex);
        }
        

    }//GEN-LAST:event_logInButtonMouseClicked

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
            java.util.logging.Logger.getLogger(LogIn.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(LogIn.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(LogIn.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(LogIn.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                l=new LogIn();
                l.setVisible(true);
            }
        });
    }
    

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JLabel jLabel4;
    private javax.swing.JPanel leftPanel;
    private javax.swing.JLabel logInButton;
    private javax.swing.JLabel password;
    private javax.swing.JPasswordField passwordTextField;
    private javax.swing.JPanel rightPanel;
    private javax.swing.JLabel userIconLeft;
    private javax.swing.JLabel userIconRight;
    private javax.swing.JTextField userNameTextField;
    private javax.swing.JLabel userPassword;
    private javax.swing.JLabel user_name;
    // End of variables declaration//GEN-END:variables

      private void setIcon() {
        setIconImage(Toolkit.getDefaultToolkit().getImage(getClass().getResource("mist_logo.png")));
    }
}
