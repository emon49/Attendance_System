

package javaapplication2;


import java.awt.Color;
import java.awt.Font;
import java.awt.Toolkit;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import static java.lang.Boolean.TRUE;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Scanner;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.JTableHeader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class attendanceSheet extends javax.swing.JFrame {

    DefaultTableModel model;
    String Cou,da;
    int colm=-1;
    public attendanceSheet() {
        initComponents();
      
        
    }
    public attendanceSheet(String Course,String date) {
        setIcon();
        initComponents();
        
        Cou=Course;
        da=date;
        showData();
       
        
        
//      JTableHeader header = jTable1.getTableHeader();
//      header.setBackground(Color.BLACK);
        
        
       jTable1.getTableHeader().setFont(new Font("Segoe UI",Font.BOLD,24));
       jTable1.getTableHeader().setOpaque(false);
//       jTable1.getTableHeader().setBackground(Color.BLACK);
      // jTable1.getTableHeader().setForeground(new Color(255,255,255));
    }

    void showData()
    {
        if(Cou.equals("JAVA"))
        {
            courseName.setText("CSE-220");
        
               
            model=(DefaultTableModel) jTable1.getModel(); 
//              model.insertRow(model.getRowCount(),new Object[]{"Faria Alam","201914012",false});
//              model.insertRow(model.getRowCount(),new Object[]{"Abdul Al Emon","201914049",false});

            try {
                Workbook workbook=new XSSFWorkbook(new FileInputStream("java_Attendance.xlsx"));
                Sheet sheet=workbook.getSheetAt(0);
                int rowCount= sheet.getPhysicalNumberOfRows();
                ////System.out.println(rowCount);
                int columCount=sheet.getRow(0).getLastCellNum();

                XSSFRow row1=(XSSFRow) sheet.getRow(0);
                //Checking whether the date is new or previous
                for(int i=2;i<columCount;i++)
                {
                    String checkdate=row1.getCell(i).toString();
                    if(da.equals(checkdate))
                    {
                        colm=i;
                        break;
                    }
                }
                
                ///System.out.println("Date: "+colm);
                
                if(colm!=-1)
                {
                    for(int i=1;i<rowCount;i++)
                    {
                        XSSFRow row=(XSSFRow) sheet.getRow(i);
                        String stuRoll=row.getCell(0).toString();
                        String stuName=row.getCell(1).toString();
                        String state=row.getCell(colm).toString();
                        if(state.equals("Pr"))
                        {
                            model.insertRow(model.getRowCount(),new Object[]{stuRoll,stuName,true});
                        }
                        else
                        {
                            model.insertRow(model.getRowCount(),new Object[]{stuRoll,stuName,false});
                        }
                        
                        

                    }
                }
                
                
                else
                {
                     for(int i=1;i<rowCount;i++)
                    {
                        XSSFRow row=(XSSFRow) sheet.getRow(i);
                        String stuRoll=row.getCell(0).toString();
                        String stuName=row.getCell(1).toString();
                        model.insertRow(model.getRowCount(),new Object[]{stuRoll,stuName,false});

                    }
                }
                

            } catch (FileNotFoundException ex) {
                Logger.getLogger(LogIn.class.getName()).log(Level.SEVERE, null, ex);
            } catch (IOException ex) {
                Logger.getLogger(LogIn.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
        
        //For Ananaya miss 
        
        else if(Cou.equals("DiscreteMath"))
        {
               courseName.setText("CSE-101");
        
               
            model=(DefaultTableModel) jTable1.getModel(); 
//              model.insertRow(model.getRowCount(),new Object[]{"Faria Alam","201914012",false});
//              model.insertRow(model.getRowCount(),new Object[]{"Abdul Al Emon","201914049",false});

            try {
                Workbook workbook=new XSSFWorkbook(new FileInputStream("Discrete.xlsx"));
                Sheet sheet=workbook.getSheetAt(0);
                int rowCount= sheet.getPhysicalNumberOfRows();
                ////System.out.println(rowCount);
                int columCount=sheet.getRow(0).getLastCellNum();

                XSSFRow row1=(XSSFRow) sheet.getRow(0);
                //Checking whether the date is new or previous
                for(int i=2;i<columCount;i++)
                {
                    String checkdate=row1.getCell(i).toString();
                    if(da.equals(checkdate))
                    {
                        colm=i;
                        break;
                    }
                }
                
                ///System.out.println("Date: "+colm);
                
                if(colm!=-1)
                {
                    for(int i=1;i<rowCount;i++)
                    {
                        XSSFRow row=(XSSFRow) sheet.getRow(i);
                        String stuRoll=row.getCell(0).toString();
                        String stuName=row.getCell(1).toString();
                        String state=row.getCell(colm).toString();
                        if(state.equals("Pr"))
                        {
                            model.insertRow(model.getRowCount(),new Object[]{stuRoll,stuName,true});
                        }
                        else
                        {
                            model.insertRow(model.getRowCount(),new Object[]{stuRoll,stuName,false});
                        }
                        
                        

                    }
                }
                
                
                else
                {
                     for(int i=1;i<rowCount;i++)
                    {
                        XSSFRow row=(XSSFRow) sheet.getRow(i);
                        String stuRoll=row.getCell(0).toString();
                        String stuName=row.getCell(1).toString();
                        model.insertRow(model.getRowCount(),new Object[]{stuRoll,stuName,false});

                    }
                }
                

                
               

                //System.out.println(Cou);
                //System.out.println(da);




            } catch (FileNotFoundException ex) {
                Logger.getLogger(LogIn.class.getName()).log(Level.SEVERE, null, ex);
            } catch (IOException ex) {
                Logger.getLogger(LogIn.class.getName()).log(Level.SEVERE, null, ex);
            }
            
        }
        

        else if(Cou.equals("DS AlGO"))
        {
              courseName.setText("CSE-215");
        
               
            model=(DefaultTableModel) jTable1.getModel(); 
//              model.insertRow(model.getRowCount(),new Object[]{"Faria Alam","201914012",false});
//              model.insertRow(model.getRowCount(),new Object[]{"Abdul Al Emon","201914049",false});

            try {
                Workbook workbook=new XSSFWorkbook(new FileInputStream("DSAlgo.xlsx"));
                Sheet sheet=workbook.getSheetAt(0);
                int rowCount= sheet.getPhysicalNumberOfRows();
                ////System.out.println(rowCount);
                int columCount=sheet.getRow(0).getLastCellNum();

                XSSFRow row1=(XSSFRow) sheet.getRow(0);
                //Checking whether the date is new or previous
                for(int i=2;i<columCount;i++)
                {
                    String checkdate=row1.getCell(i).toString();
                    if(da.equals(checkdate))
                    {
                        colm=i;
                        break;
                    }
                }
                
                ///System.out.println("Date: "+colm);
                
                if(colm!=-1)
                {
                    for(int i=1;i<rowCount;i++)
                    {
                        XSSFRow row=(XSSFRow) sheet.getRow(i);
                        String stuRoll=row.getCell(0).toString();
                        String stuName=row.getCell(1).toString();
                        String state=row.getCell(colm).toString();
                        if(state.equals("Pr"))
                        {
                            model.insertRow(model.getRowCount(),new Object[]{stuRoll,stuName,true});
                        }
                        else
                        {
                            model.insertRow(model.getRowCount(),new Object[]{stuRoll,stuName,false});
                        }
                        
                        

                    }
                }
                
                
                else
                {
                     for(int i=1;i<rowCount;i++)
                    {
                        XSSFRow row=(XSSFRow) sheet.getRow(i);
                        String stuRoll=row.getCell(0).toString();
                        String stuName=row.getCell(1).toString();
                        model.insertRow(model.getRowCount(),new Object[]{stuRoll,stuName,false});

                    }
                }
                

                
               

                //System.out.println(Cou);
                //System.out.println(da);




            } catch (FileNotFoundException ex) {
                Logger.getLogger(LogIn.class.getName()).log(Level.SEVERE, null, ex);
            } catch (IOException ex) {
                Logger.getLogger(LogIn.class.getName()).log(Level.SEVERE, null, ex);
            }
            
        }
         
    }
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        attendanceForm = new javax.swing.JPanel();
        jScrollPane1 = new javax.swing.JScrollPane();
        jTable1 = new javax.swing.JTable();
        courseName = new javax.swing.JLabel();
        courseCode = new javax.swing.JLabel();
        saveButton = new javax.swing.JLabel();
        chooseFileButton = new javax.swing.JLabel();
        fileChooserField = new javax.swing.JTextField();
        logOutAttendance = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        attendanceForm.setBackground(new java.awt.Color(44, 44, 84));
        attendanceForm.setForeground(new java.awt.Color(255, 255, 255));
        attendanceForm.setMinimumSize(new java.awt.Dimension(868, 506));
        attendanceForm.setPreferredSize(new java.awt.Dimension(876, 748));
        attendanceForm.setLayout(null);

        jTable1.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jTable1.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "ID", "Name", "Present"
            }
        ) {
            Class[] types = new Class [] {
                java.lang.Object.class, java.lang.Object.class, java.lang.Boolean.class
            };
            boolean[] canEdit = new boolean [] {
                false, false, true
            };

            public Class getColumnClass(int columnIndex) {
                return types [columnIndex];
            }

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        jTable1.setFocusable(false);
        jTable1.setIntercellSpacing(new java.awt.Dimension(0, 0));
        jTable1.setRowHeight(25);
        jTable1.getTableHeader().setReorderingAllowed(false);
        jScrollPane1.setViewportView(jTable1);

        attendanceForm.add(jScrollPane1);
        jScrollPane1.setBounds(170, 100, 620, 530);

        courseName.setFont(new java.awt.Font("Segoe UI", 1, 28)); // NOI18N
        courseName.setForeground(new java.awt.Color(255, 255, 255));
        attendanceForm.add(courseName);
        courseName.setBounds(430, 40, 250, 45);

        courseCode.setFont(new java.awt.Font("Segoe UI", 1, 28)); // NOI18N
        courseCode.setForeground(new java.awt.Color(255, 255, 255));
        courseCode.setText("Course Code:");
        attendanceForm.add(courseCode);
        courseCode.setBounds(240, 40, 180, 45);

        saveButton.setIcon(new javax.swing.ImageIcon(getClass().getResource("/javaapplication2/saveBtn1.png"))); // NOI18N
        saveButton.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        saveButton.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                saveButtonMouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                saveButtonMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                saveButtonMouseExited(evt);
            }
        });
        attendanceForm.add(saveButton);
        saveButton.setBounds(400, 780, 120, 50);

        chooseFileButton.setIcon(new javax.swing.ImageIcon(getClass().getResource("/javaapplication2/choose_file.png"))); // NOI18N
        chooseFileButton.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                chooseFileButtonMouseClicked(evt);
            }
        });
        attendanceForm.add(chooseFileButton);
        chooseFileButton.setBounds(570, 690, 140, 30);
        attendanceForm.add(fileChooserField);
        fileChooserField.setBounds(280, 682, 270, 40);

        logOutAttendance.setIcon(new javax.swing.ImageIcon(getClass().getResource("/javaapplication2/logoutAttendance.png"))); // NOI18N
        logOutAttendance.setCursor(new java.awt.Cursor(java.awt.Cursor.HAND_CURSOR));
        logOutAttendance.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                logOutAttendanceMouseClicked(evt);
            }
        });
        attendanceForm.add(logOutAttendance);
        logOutAttendance.setBounds(740, 20, 110, 40);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(attendanceForm, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(attendanceForm, javax.swing.GroupLayout.DEFAULT_SIZE, 865, Short.MAX_VALUE)
        );

        pack();
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void saveButtonMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_saveButtonMouseEntered
        saveButton.setIcon(new javax.swing.ImageIcon(getClass().getResource("./saveBtn2.png")));
    }//GEN-LAST:event_saveButtonMouseEntered

    private void saveButtonMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_saveButtonMouseExited
        saveButton.setIcon(new javax.swing.ImageIcon(getClass().getResource("./saveBtn1.png")));
    }//GEN-LAST:event_saveButtonMouseExited

    private void saveButtonMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_saveButtonMouseClicked
        int pr=0,ab=0;
        int rowCount=jTable1.getRowCount();
        int[] intArray = new int[rowCount];
        Workbook workbook = null;
        //java starts
       if(Cou.equals("JAVA")){ 
        
         try {
             workbook = new XSSFWorkbook(new FileInputStream("java_Attendance.xlsx"));
         } catch (FileNotFoundException ex) {
             Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
         } catch (IOException ex) {
             Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
         }
        Sheet sheet=workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();
        Row row = rowIterator.next();
        
        if(colm!=-1)
        {
             for(int i=0;i<rowCount;i++)  
            {
               Boolean chk= ((Boolean)jTable1.getValueAt(i, 2)).booleanValue();
               if(chk)
              {       pr++;  
                    //columCount=sheet.getRow(i+1).getLastCellNum();                    
                    row = rowIterator.next();
                    //Cell cell =row.createCell(columCount);
                    Cell cell=row.getCell(colm);
                    cell.setCellValue("Pr");
                    try {                
                        FileOutputStream out = new  FileOutputStream(new File("java_Attendance.xlsx"));
                        try {
                            workbook.write(out);
                           
                            out.close();
                        } catch (IOException ex) {
                            Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                        }
                        
                    } catch (FileNotFoundException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (IOException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } 

                    
                
                }
            else
            {   ab++;
               // columCount=sheet.getRow(i+1).getLastCellNum();                    
                    row = rowIterator.next();
                    Cell cell =row.getCell(colm);
                    cell.setCellValue("Ab");
                    try {                
                        FileOutputStream out = new  FileOutputStream(new File("java_Attendance.xlsx"));
                        try {
                            workbook.write(out);
                            out.close();
                        } catch (IOException ex) {
                            Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                        }
                        
                    } catch (FileNotFoundException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (IOException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } 

            }
          }
       
       }
        
        //new date 
        else
        {
        
        
/*        Setting Date Value            */
        int columCount=sheet.getRow(0).getLastCellNum();  
        row.createCell(columCount).setCellValue(da);
                
                
                
/*        Reading value from JTable & Set them to Excel     */

      
                
        for(int i=0;i<rowCount;i++)  
        {
            Boolean chk= ((Boolean)jTable1.getValueAt(i, 2)).booleanValue();
            if(chk)
            {       pr++;  
                    columCount=sheet.getRow(i+1).getLastCellNum();                    
                    row = rowIterator.next();
                    Cell cell =row.createCell(columCount);
                    cell.setCellValue("Pr");
                    try {                
                        FileOutputStream out = new  FileOutputStream(new File("java_Attendance.xlsx"));
                        try {
                            workbook.write(out);
                           
                            out.close();
                        } catch (IOException ex) {
                            Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                        }
                        
                    } catch (FileNotFoundException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (IOException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } 

                    
                
                }
            else
            {   ab++;
                columCount=sheet.getRow(i+1).getLastCellNum();                    
                    row = rowIterator.next();
                    Cell cell =row.createCell(columCount);
                    cell.setCellValue("Ab");
                    try {                
                        FileOutputStream out = new  FileOutputStream(new File("java_Attendance.xlsx"));
                        try {
                            workbook.write(out);
                            out.close();
                        } catch (IOException ex) {
                            Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                        }
                        
                    } catch (FileNotFoundException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (IOException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } 

            }
        }
       
        
        
        } //new datw else ends 
        
       }
       
       
       
       
       
       
       //Discrete math
       
       else if(Cou.equals("DiscreteMath"))
       {
           try {
             workbook = new XSSFWorkbook(new FileInputStream("Discrete.xlsx"));
         } catch (FileNotFoundException ex) {
             Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
         } catch (IOException ex) {
             Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
         }
        Sheet sheet=workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();
        Row row = rowIterator.next();
        
        if(colm!=-1)
        {
             for(int i=0;i<rowCount;i++)  
            {
               Boolean chk= ((Boolean)jTable1.getValueAt(i, 2)).booleanValue();
               if(chk)
              {       pr++;  
                    //columCount=sheet.getRow(i+1).getLastCellNum();                    
                    row = rowIterator.next();
                    //Cell cell =row.createCell(columCount);
                    Cell cell=row.getCell(colm);
                    cell.setCellValue("Pr");
                    try {                
                        FileOutputStream out = new  FileOutputStream(new File("Discrete.xlsx"));
                        try {
                            workbook.write(out);
                           
                            out.close();
                        } catch (IOException ex) {
                            Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                        }
                        
                    } catch (FileNotFoundException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (IOException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } 

                    
                
                }
            else
            {   ab++;
               // columCount=sheet.getRow(i+1).getLastCellNum();                    
                    row = rowIterator.next();
                    Cell cell =row.getCell(colm);
                    cell.setCellValue("Ab");
                    try {                
                        FileOutputStream out = new  FileOutputStream(new File("Discrete.xlsx"));
                        try {
                            workbook.write(out);
                            out.close();
                        } catch (IOException ex) {
                            Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                        }
                        
                    } catch (FileNotFoundException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (IOException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } 

            }
          }
       
       }
        
        //new date 
        else
        {
        
        
/*        Setting Date Value            */
        int columCount=sheet.getRow(0).getLastCellNum();  
        row.createCell(columCount).setCellValue(da);
                
                
                
/*        Reading value from JTable & Set them to Excel     */

      
                
        for(int i=0;i<rowCount;i++)  
        {
            Boolean chk= ((Boolean)jTable1.getValueAt(i, 2)).booleanValue();
            if(chk)
            {       pr++;  
                    columCount=sheet.getRow(i+1).getLastCellNum();                    
                    row = rowIterator.next();
                    Cell cell =row.createCell(columCount);
                    cell.setCellValue("Pr");
                    try {                
                        FileOutputStream out = new  FileOutputStream(new File("Discrete.xlsx"));
                        try {
                            workbook.write(out);
                           
                            out.close();
                        } catch (IOException ex) {
                            Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                        }
                        
                    } catch (FileNotFoundException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (IOException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } 

                    
                
                }
            else
            {   ab++;
                columCount=sheet.getRow(i+1).getLastCellNum();                    
                    row = rowIterator.next();
                    Cell cell =row.createCell(columCount);
                    cell.setCellValue("Ab");
                    try {                
                        FileOutputStream out = new  FileOutputStream(new File("Discrete.xlsx"));
                        try {
                            workbook.write(out);
                            out.close();
                        } catch (IOException ex) {
                            Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                        }
                        
                    } catch (FileNotFoundException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (IOException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } 

            }
        }
       
        
        
        } //new datw else ends 
       }//Discrete math ends
       
       
       
       //DSALGO Starts
       
       else if(Cou.equals("DS AlGO"))
       {
           System.out.println("Helllo ds algo");
           try {
             workbook = new XSSFWorkbook(new FileInputStream("DSAlgo.xlsx"));
         } catch (FileNotFoundException ex) {
             Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
         } catch (IOException ex) {
             Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
         }
        Sheet sheet=workbook.getSheetAt(0);
        Iterator<Row> rowIterator = sheet.iterator();
        Row row = rowIterator.next();
        
        if(colm!=-1)
        {
             for(int i=0;i<rowCount;i++)  
            {
               Boolean chk= ((Boolean)jTable1.getValueAt(i, 2)).booleanValue();
               if(chk)
              {       pr++;  
                    //columCount=sheet.getRow(i+1).getLastCellNum();                    
                    row = rowIterator.next();
                    //Cell cell =row.createCell(columCount);
                    Cell cell=row.getCell(colm);
                    cell.setCellValue("Pr");
                    try {                
                        FileOutputStream out = new  FileOutputStream(new File("DSAlgo.xlsx"));
                        try {
                            workbook.write(out);
                           
                            out.close();
                        } catch (IOException ex) {
                            Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                        }
                        
                    } catch (FileNotFoundException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (IOException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } 

                    
                
                }
            else
            {   ab++;
               // columCount=sheet.getRow(i+1).getLastCellNum();                    
                    row = rowIterator.next();
                    Cell cell =row.getCell(colm);
                    cell.setCellValue("Ab");
                    try {                
                        FileOutputStream out = new  FileOutputStream(new File("DSAlgo.xlsx"));
                        try {
                            workbook.write(out);
                            out.close();
                        } catch (IOException ex) {
                            Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                        }
                        
                    } catch (FileNotFoundException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (IOException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } 

            }
          }
       
       }
        
        //new date 
        else
        {
        
        
/*        Setting Date Value            */
        int columCount=sheet.getRow(0).getLastCellNum();  
        row.createCell(columCount).setCellValue(da);
                
                
                
/*        Reading value from JTable & Set them to Excel     */

      
                
        for(int i=0;i<rowCount;i++)  
        {
            Boolean chk= ((Boolean)jTable1.getValueAt(i, 2)).booleanValue();
            if(chk)
            {       pr++;  
                    columCount=sheet.getRow(i+1).getLastCellNum();                    
                    row = rowIterator.next();
                    Cell cell =row.createCell(columCount);
                    cell.setCellValue("Pr");
                    try {                
                        FileOutputStream out = new  FileOutputStream(new File("DSAlgo.xlsx"));
                        try {
                            workbook.write(out);
                           
                            out.close();
                        } catch (IOException ex) {
                            Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                        }
                        
                    } catch (FileNotFoundException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (IOException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } 

                    
                
                }
            else
            {   ab++;
                columCount=sheet.getRow(i+1).getLastCellNum();                    
                    row = rowIterator.next();
                    Cell cell =row.createCell(columCount);
                    cell.setCellValue("Ab");
                    try {                
                        FileOutputStream out = new  FileOutputStream(new File("DSAlgo.xlsx"));
                        try {
                            workbook.write(out);
                            out.close();
                        } catch (IOException ex) {
                            Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                        }
                        
                    } catch (FileNotFoundException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } catch (IOException ex) {
                        Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
                    } 

            }
        }
       
        
        
        } //new datw else ends 
       }
       
       
       
        
        summaryPage su=new summaryPage(Cou,da,pr,ab);
        this.setVisible(false);
        su.setVisible(true);
        
    }//GEN-LAST:event_saveButtonMouseClicked

    private void chooseFileButtonMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_chooseFileButtonMouseClicked
        
        JFileChooser chooser=new JFileChooser();
        chooser.showOpenDialog(null);
        File f=chooser.getSelectedFile();
        String filename=f.getAbsolutePath();
        fileChooserField.setText(filename);
        
        
         ArrayList<String> rollList=new ArrayList<String>();
        
        FileInputStream fis;
        
         try {
            fis = new FileInputStream(filename);
            Scanner sc=new Scanner(fis);
             String s="";
         while(sc.hasNextLine())
        {
            s="";
             String roll="";
            s=s.concat(sc.nextLine());
            int len=s.length();
            for(int i=len-1;i>=len-9;i--)
            {
                roll=roll+s.charAt(i);
                
            }
             StringBuilder input1 = new StringBuilder(); 

             input1.append(roll);
             input1 = input1.reverse(); 
             String sroll=input1.toString();
             rollList.add(sroll);
            
            //System.out.println(input1);
        }
         
         int r_count=jTable1.getRowCount();
         
         int len=rollList.size();
        
         for(int r=0;r<len;r++) //Array list from saved chat
         {
             for(int l=0;l<r_count;l++) //For settng in model
             {
                 if(rollList.get(r).equals(model.getValueAt(l, 0).toString())){
                     model.setValueAt(TRUE,l,2);
                 }
             }
         }
     
        } catch (FileNotFoundException ex) {
            Logger.getLogger(attendanceSheet.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_chooseFileButtonMouseClicked

    private void logOutAttendanceMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_logOutAttendanceMouseClicked
        this.setVisible(false);
       LogIn ln=new LogIn();
       ln.setVisible(true);
    }//GEN-LAST:event_logOutAttendanceMouseClicked

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
            java.util.logging.Logger.getLogger(attendanceSheet.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(attendanceSheet.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(attendanceSheet.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(attendanceSheet.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new attendanceSheet().setVisible(true);
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JPanel attendanceForm;
    private javax.swing.JLabel chooseFileButton;
    private javax.swing.JLabel courseCode;
    private javax.swing.JLabel courseName;
    private javax.swing.JTextField fileChooserField;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTable jTable1;
    private javax.swing.JLabel logOutAttendance;
    private javax.swing.JLabel saveButton;
    // End of variables declaration//GEN-END:variables
 private void setIcon() {
        setIconImage(Toolkit.getDefaultToolkit().getImage(getClass().getResource("mist_logo.png")));
    }
}
