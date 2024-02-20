
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.*;
import java.io.*;
import java.util.*;
import java.sql.*;
import java.util.Vector;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import java.awt.Desktop;
import static org.apache.commons.math3.fitting.leastsquares.LeastSquaresFactory.model;

/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */

/**
 *
 * @author Arnie D
 */

public class Form extends javax.swing.JFrame {

    /**
     * Creates new form Form
     */
    public Form() {
        initComponents();
        Connect();
    }
    
    Connection con; 
    PreparedStatement pst;
    ResultSet rs;
    
    public void Connect () {
        try {
            Class.forName("com.mysql.cj.jdbc.Driver");
            con = DriverManager.getConnection("jdbc:mysql://localhost:3306/jdc","root","admin");
        } catch (ClassNotFoundException | SQLException ex) {
            Logger.getLogger(Form.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    public void Import(){

        File excelFile;
        FileInputStream excelFIS = null;
        BufferedInputStream excelBIS = null;
        XSSFWorkbook excelJTableImport = null;

        String defaultCurrentDirectoryPath = "C:\\Users\\Arnie D\\Downloads\\Arnie\\OJT\\IT DEPT";
        JFileChooser excelFileChooser = new JFileChooser(defaultCurrentDirectoryPath);

        FileNameExtensionFilter fnef = new FileNameExtensionFilter("EXCEL FILES", "xls", "xlsx", "xlsm");
        excelFileChooser.setFileFilter(fnef);
        int excelChooser = excelFileChooser.showOpenDialog(null);

        if (excelChooser == JFileChooser.APPROVE_OPTION) {
            try {
                excelFile = excelFileChooser.getSelectedFile();
                excelFIS = new FileInputStream(excelFile);
                excelBIS = new BufferedInputStream(excelFIS);
                excelJTableImport = new XSSFWorkbook(excelBIS);
                XSSFSheet excelSheet = excelJTableImport.getSheetAt(0);

                for (int row = 1; row <= excelSheet.getLastRowNum(); row++) {
                    XSSFRow excelRow = excelSheet.getRow(row);

                    XSSFCell eDept = excelRow.getCell(0);
                    XSSFCell eAct = excelRow.getCell(1);
                    XSSFCell eDB = excelRow.getCell(2);
                    XSSFCell eUCode = excelRow.getCell(3);
                    XSSFCell eName = excelRow.getCell(4);
                    XSSFCell eLicense = excelRow.getCell(5);
                    XSSFCell eRemarks = excelRow.getCell(6);

                    DefaultTableModel df = (DefaultTableModel) table.getModel();
                    df.addRow(new Object[]{eDept, eAct, eDB, eUCode, eName, eLicense, eRemarks});

                }
                JOptionPane.showMessageDialog(null, "Imported Successfully");
                

                

            } catch (FileNotFoundException ex) {
                JOptionPane.showMessageDialog(null, ex.getMessage());
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(null, ex.getMessage());
            } finally {
                try {
                    if (excelFIS != null) {
                        excelFIS.close();
                    }
                    if (excelBIS != null) {
                        excelBIS.close();
                    }
                    if (excelJTableImport != null) {
                        excelJTableImport.close();
                    }
                } catch (IOException ioException) {

                }
            }

        }
    }
        public void search_userCode(){
        try {
            String code = user_code.getText();
            pst = con.prepareStatement("Select * FROM user WHERE user_code=?");
            pst.setString(1, code);
            rs = pst.executeQuery();
            DefaultTableModel model = (DefaultTableModel) table.getModel();
            model.setRowCount(0);
            String search = "false";
            while (rs.next()) {
                search = "true";
                String dept = rs.getString("department");
                String act = rs.getString("activity");
                String db = rs.getString("db");
                String ucode = rs.getString("user_code");
                String name = rs.getString("name");
                String license = rs.getString("license");
                String remarks = rs.getString("remarls");

                model.addRow(new Object[]{dept, act, db, ucode, name, license, remarks});

            }
           if(search == "false" && !rs.next()){
                JOptionPane.showMessageDialog(null, "No data found");
                isEmpty();
            }

        } catch (Exception e) {
            System.out.print(e);

        }
    }
         public void search_act(){
        try {
            String code = act1.getText();
            pst = con.prepareStatement("Select * FROM user WHERE activity=?");
            pst.setString(1, code);
            rs = pst.executeQuery();
            DefaultTableModel model = (DefaultTableModel) table.getModel();
            model.setRowCount(0);
            String search = "false";
            while (rs.next()) {
                search = "true";
                String dept = rs.getString("department");
                String act = rs.getString("activity");
                String db = rs.getString("db");
                String ucode = rs.getString("user_code");
                String name = rs.getString("name");
                String license = rs.getString("license");
                String remarks = rs.getString("remarls");

                model.addRow(new Object[]{dept, act, db, ucode, name, license, remarks});

            }
           if(search == "false" && !rs.next()){
                JOptionPane.showMessageDialog(null, "No data found");
                isEmpty();
            }

        } catch (Exception e) {
            System.out.print(e);

        }
    }

    public void search_name(){
        try {
            String n1 = name.getText();
            pst = con.prepareStatement("Select * FROM user WHERE name=?");
            pst.setString(1, n1);
            rs = pst.executeQuery();
            DefaultTableModel model = (DefaultTableModel) table.getModel();
            model.setRowCount(0);
            String search = "false";
            while (rs.next()) {
                search = "true";
                String dept = rs.getString("department");
                String act = rs.getString("activity");
                String db = rs.getString("db");
                String ucode = rs.getString("user_code");
                String name = rs.getString("name");
                String license = rs.getString("license");
                String remarks = rs.getString("remarls");

                model.addRow(new Object[]{dept, act, db, ucode, name, license, remarks});

            }
             if(search == "false" && !rs.next()){
                 
                JOptionPane.showMessageDialog(null, "No data found");
                isEmpty();
            }

        } catch (Exception e) {
            System.out.print(e);

        }
    }
    public void search_dep(){
        try {
            String n1 = dept.getText();
            pst = con.prepareStatement("Select * FROM user WHERE department=?");
            pst.setString(1, n1);
            rs = pst.executeQuery();
            DefaultTableModel model = (DefaultTableModel) table.getModel();
            model.setRowCount(0);
            String search = "false";
            while (rs.next()) {
                search = "true";
                String dept = rs.getString("department");
                String act = rs.getString("activity");
                String db = rs.getString("db");
                String ucode = rs.getString("user_code");
                String name = rs.getString("name");
                String license = rs.getString("license");
                String remarks = rs.getString("remarls");

                model.addRow(new Object[]{dept, act, db, ucode, name, license, remarks});

            }
             if(search == "false" && !rs.next()){
                JOptionPane.showMessageDialog(null, "No data found");
                isEmpty();
            }

        } catch (Exception e) {
            System.out.print(e);

        }
    }
     public void search_license(){
        try {
            String n1 = LType.getSelectedItem().toString();
            pst = con.prepareStatement("Select * FROM user WHERE license=?");
            pst.setString(1, n1);
            rs = pst.executeQuery();
            DefaultTableModel model = (DefaultTableModel) table.getModel();
            model.setRowCount(0);
            String search = "false";
            while (rs.next()) {
                search = "true";
                String dept = rs.getString("department");
                String act = rs.getString("activity");
                String db = rs.getString("db");
                String ucode = rs.getString("user_code");
                String name = rs.getString("name");
                String license = rs.getString("license");
                String remarks = rs.getString("remarls");

                model.addRow(new Object[]{dept, act, db, ucode, name, license, remarks});

            }
             if(search == "false" && !rs.next()){
                     JOptionPane.showMessageDialog(null, "No data found");
                isEmpty();
                
            }

        } catch (Exception e) {
            System.out.print(e);

        }
    }
    public void isEmpty() {
        try {
            pst = con.prepareStatement("Select * from user");
            rs = pst.executeQuery();
            DefaultTableModel model = (DefaultTableModel) table.getModel();
            Object[] row;
            while(rs.next()){
                row = new Object[7];
                row[0] = rs.getString(1);
                row[1] = rs.getString(2);
                row[2] = rs.getString(3);
                row[3] = rs.getString(4);
                row[4] = rs.getString(5);
                row[5] = rs.getString(6);
                row[6] = rs.getString(7);
                
                model.addRow(row);
            }
            model.fireTableDataChanged();
        } catch (Exception ex) {
            ex.printStackTrace();
        }finally {
            try {
                if (rs != null) {
                    rs.close();
                }
                if (pst != null) {
                    pst.close();
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    

    }
    public void openFile(String file){
        try{
            File path = new File(file);
            Desktop.getDesktop().open(path);
        }catch(IOException ioe){
            System.out.println(ioe);
        }
    }

    
    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        jScrollPane1 = new javax.swing.JScrollPane();
        table = new javax.swing.JTable();
        jButton1 = new javax.swing.JButton();
        SaveDB = new javax.swing.JButton();
        jLabel1 = new javax.swing.JLabel();
        user_code = new javax.swing.JTextField();
        jLabel2 = new javax.swing.JLabel();
        name = new javax.swing.JTextField();
        jLabel3 = new javax.swing.JLabel();
        dept = new javax.swing.JTextField();
        jLabel4 = new javax.swing.JLabel();
        LType = new javax.swing.JComboBox<>();
        export = new javax.swing.JButton();
        SaveDB1 = new javax.swing.JButton();
        jButton2 = new javax.swing.JButton();
        act1 = new javax.swing.JTextField();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setLocation(new java.awt.Point(0, 0));
        setMaximumSize(new java.awt.Dimension(1650, 1080));

        jPanel1.setBackground(new java.awt.Color(255, 255, 255));
        jPanel1.setInheritsPopupMenu(true);

        table.setAutoCreateRowSorter(true);
        table.setFont(new java.awt.Font("Microsoft JhengHei", 0, 14)); // NOI18N
        table.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "DEPARTMENT", "ACTIVITY", "DATABASE", "USER CODE", "NAME", "LICENSE TYPE", "REMARKS"
            }
        ) {
            Class[] types = new Class [] {
                java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class
            };
            boolean[] canEdit = new boolean [] {
                false, false, false, false, false, false, false
            };

            public Class getColumnClass(int columnIndex) {
                return types [columnIndex];
            }

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        table.setGridColor(new java.awt.Color(255, 255, 255));
        table.setSelectionBackground(new java.awt.Color(255, 0, 0));
        table.setSelectionForeground(new java.awt.Color(255, 255, 255));
        table.getTableHeader().setReorderingAllowed(false);
        jScrollPane1.setViewportView(table);

        jButton1.setBackground(new java.awt.Color(0, 102, 51));
        jButton1.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jButton1.setForeground(new java.awt.Color(255, 255, 255));
        jButton1.setText("IMPORT EXCEL FILE");
        jButton1.setMaximumSize(new java.awt.Dimension(84, 32));
        jButton1.setMinimumSize(new java.awt.Dimension(84, 32));
        jButton1.setPreferredSize(new java.awt.Dimension(84, 32));
        jButton1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton1ActionPerformed(evt);
            }
        });

        SaveDB.setBackground(new java.awt.Color(0, 102, 51));
        SaveDB.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        SaveDB.setForeground(new java.awt.Color(255, 255, 255));
        SaveDB.setText("SAVE");
        SaveDB.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                SaveDBActionPerformed(evt);
            }
        });

        jLabel1.setText("User Code:");

        user_code.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                user_codeActionPerformed(evt);
            }
        });

        jLabel2.setText("Name:");

        jLabel3.setText("Department:");

        jLabel4.setText("License Type:");

        LType.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "   ", "Item 1", "Item 2", "Item 3", "Item 4", "1515" }));

        export.setBackground(new java.awt.Color(0, 102, 51));
        export.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        export.setForeground(new java.awt.Color(255, 255, 255));
        export.setText("EXPORT TO EXCEL FILE");
        export.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                exportActionPerformed(evt);
            }
        });

        SaveDB1.setBackground(new java.awt.Color(0, 102, 51));
        SaveDB1.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        SaveDB1.setForeground(new java.awt.Color(255, 255, 255));
        SaveDB1.setText("Search");
        SaveDB1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                SaveDB1ActionPerformed(evt);
            }
        });

        jButton2.setText("jButton2");
        jButton2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton2ActionPerformed(evt);
            }
        });

        act1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                act1ActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addGap(111, 111, 111)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 245, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(138, 138, 138)
                        .addComponent(jLabel4)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(LType, javax.swing.GroupLayout.PREFERRED_SIZE, 200, javax.swing.GroupLayout.PREFERRED_SIZE))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                            .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 176, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 176, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGroup(jPanel1Layout.createSequentialGroup()
                                    .addGap(6, 6, 6)
                                    .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                                        .addComponent(name, javax.swing.GroupLayout.PREFERRED_SIZE, 276, javax.swing.GroupLayout.PREFERRED_SIZE)
                                        .addComponent(dept, javax.swing.GroupLayout.PREFERRED_SIZE, 276, javax.swing.GroupLayout.PREFERRED_SIZE))))
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGap(6, 6, 6)
                                .addComponent(user_code, javax.swing.GroupLayout.PREFERRED_SIZE, 276, javax.swing.GroupLayout.PREFERRED_SIZE)))
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                                .addComponent(jButton2))
                            .addGroup(jPanel1Layout.createSequentialGroup()
                                .addGap(77, 77, 77)
                                .addComponent(act1, javax.swing.GroupLayout.PREFERRED_SIZE, 276, javax.swing.GroupLayout.PREFERRED_SIZE)
                                .addGap(0, 0, Short.MAX_VALUE)))))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 246, Short.MAX_VALUE)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(SaveDB, javax.swing.GroupLayout.PREFERRED_SIZE, 197, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                        .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                            .addComponent(jButton1, javax.swing.GroupLayout.DEFAULT_SIZE, 197, Short.MAX_VALUE)
                            .addComponent(SaveDB1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                        .addComponent(export, javax.swing.GroupLayout.PREFERRED_SIZE, 197, javax.swing.GroupLayout.PREFERRED_SIZE)))
                .addGap(44, 44, 44))
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(jScrollPane1)
                .addContainerGap())
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(javax.swing.GroupLayout.Alignment.TRAILING, jPanel1Layout.createSequentialGroup()
                .addGap(19, 19, 19)
                .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jLabel1, javax.swing.GroupLayout.PREFERRED_SIZE, 33, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jLabel4, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                            .addComponent(LType))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(user_code, javax.swing.GroupLayout.PREFERRED_SIZE, 25, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                        .addComponent(jLabel2, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(4, 4, 4)
                        .addComponent(name, javax.swing.GroupLayout.PREFERRED_SIZE, 28, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(jLabel3, javax.swing.GroupLayout.PREFERRED_SIZE, 30, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(dept, javax.swing.GroupLayout.PREFERRED_SIZE, 31, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(jButton2)))
                    .addGroup(jPanel1Layout.createSequentialGroup()
                        .addComponent(SaveDB1, javax.swing.GroupLayout.PREFERRED_SIZE, 46, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addGroup(jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                            .addComponent(jButton1, javax.swing.GroupLayout.PREFERRED_SIZE, 50, javax.swing.GroupLayout.PREFERRED_SIZE)
                            .addComponent(act1, javax.swing.GroupLayout.PREFERRED_SIZE, 25, javax.swing.GroupLayout.PREFERRED_SIZE))
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(export, javax.swing.GroupLayout.PREFERRED_SIZE, 50, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                        .addComponent(SaveDB)
                        .addGap(0, 0, Short.MAX_VALUE)))
                .addGap(47, 47, 47)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 393, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(52, 52, 52))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.Alignment.TRAILING, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );

        pack();
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        int RowCount = table.getRowCount();
        if (RowCount != 0) {
            DefaultTableModel df = (DefaultTableModel) table.getModel();
            df.setRowCount(0);
            RowCount = 0;
            if (RowCount == 0) {
                Import();
            }
        } else {
            Import();
        }
        
  
    }//GEN-LAST:event_jButton1ActionPerformed
   
    private void SaveDBActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_SaveDBActionPerformed
        int columnIndex = 3; 
        DefaultTableModel model = (DefaultTableModel) table.getModel();
        int rowCount = model.getRowCount();
        int saved = 0;
        int dup = 1;
        if(rowCount != 0){
        for (int row = 0; row < rowCount; row++) {
            Object value = model.getValueAt(row, columnIndex);
            try{
            
                String duplicate = "SELECT * FROM user WHERE user_code = '" + value + "'";
                pst = con.prepareStatement(duplicate);
                rs = pst.executeQuery(duplicate);
                
                if (rs.next()) {
                    if (dup == 1) {
                        dup=0;
                        JOptionPane.showMessageDialog(this, "Some User Code already used ");
                    }

                }
         
           else if (!rs.next()){
                   
                    int rowCount1 = table.getRowCount();
                    int columnCount = table.getColumnCount();
                    for (int row1 = 0; row1 < rowCount1; row1++) {
                       
                            Object val = table.getValueAt(row1, 3);
                            Object dept = table.getValueAt(row1, 0);
                            Object act = table.getValueAt(row1, 1);
                            Object db = table.getValueAt(row1, 2);
                            Object user_code = table.getValueAt(row1, 3);
                            Object name = table.getValueAt(row1, 4);
                            Object license = table.getValueAt(row1, 5);
                            Object remarks = table.getValueAt(row1, 6);
                            
                           
                            if (val == value) {
                                
                                pst = con.prepareStatement("Insert into user(department,activity,db,user_code,name,license,remarls) values (?,?,?,?,?,?,?)");
                                pst.setString(1, dept.toString());
                                pst.setString(2, act.toString());
                                pst.setString(3, db.toString());
                                pst.setString(4, user_code.toString());
                                pst.setString(5, name.toString());
                                pst.setString(6, license.toString());
                                pst.setString(7, remarks.toString());

                                saved = pst.executeUpdate();
                              
                            }

                        }

                    }

                } catch (SQLException ex) {
                    Logger.getLogger(Form.class.getName()).log(Level.SEVERE, null, ex);
                }

            }
            if (saved == 1) {
                JOptionPane.showMessageDialog(null, "SAVED!");
            }
        } else {
            JOptionPane.showMessageDialog(null, "You need to import files");
        }
    }//GEN-LAST:event_SaveDBActionPerformed

    private void user_codeActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_user_codeActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_user_codeActionPerformed
 
        
    private void exportActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_exportActionPerformed
    try{
        JFileChooser jFileChooser = new JFileChooser();
        jFileChooser.showSaveDialog(this);
        File saveFile = jFileChooser.getSelectedFile();
        if(saveFile !=null){
            saveFile = new File(saveFile.toString()+".xlsx");
            Workbook wb = new XSSFWorkbook();
            Sheet sheet = wb.createSheet("customer");
            
            Row rowCol = sheet.createRow(0);
            for (int c=0; c< table.getColumnCount();c++){
                Cell cell = rowCol.createCell(c);
                cell.setCellValue(table.getColumnName(c));
            }
            for(int r=0;r<table.getRowCount();r++){
                Row row = sheet.createRow(r);
                for(int rc=0;rc<table.getColumnCount();rc++){
                    Cell cell = row.createCell(rc);
                    if(table.getValueAt(r, rc)!=null){
                        cell.setCellValue(table.getValueAt(r,rc).toString());
                    }
                }
            }
            FileOutputStream out = new FileOutputStream(new File(saveFile.toString()));
            wb.write(out);
            wb.close();
            out.close();
            openFile(saveFile.toString());
        }else{
            JOptionPane.showMessageDialog(null,"Error!");
        }
        
    }catch(FileNotFoundException e){
        System.out.println(e);
    }catch(IOException io){
        System.out.println(io);
    }
    }//GEN-LAST:event_exportActionPerformed

    private void jButton2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton2ActionPerformed
        this.dispose();
        Employee emp = new Employee();
        emp.setVisible(true);
    }//GEN-LAST:event_jButton2ActionPerformed

    private void SaveDB1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_SaveDB1ActionPerformed
        String code= user_code.getText();
        String n1 = name.getText();
        String d1 = dept.getText();
        String lt=  LType.getSelectedItem().toString();
        if (code.equals("") && (n1.equals("") && (d1.equals("") &&(lt.isEmpty())))) {
            {
                DefaultTableModel df = (DefaultTableModel) table.getModel();
                df.setRowCount(0);
                isEmpty();
            }

        }
        else{
            if(!code.isEmpty()){
                search_userCode();
            }
            else if(!n1.isEmpty()){
                search_name();
            }
            else if(!d1.isEmpty()){
                search_dep();
            }
            else if(!lt.isEmpty()){
                search_license();
            }
            user_code.setText("");
            name.setText("");
            dept.setText("");
            LType.setSelectedIndex(0);
        }

    }//GEN-LAST:event_SaveDB1ActionPerformed

    private void act1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_act1ActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_act1ActionPerformed
 
   
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
            java.util.logging.Logger.getLogger(Form.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Form.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Form.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Form.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new Form().setVisible(true);
                 
                
            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JComboBox<String> LType;
    private javax.swing.JButton SaveDB;
    private javax.swing.JButton SaveDB1;
    private javax.swing.JTextField act1;
    private javax.swing.JTextField dept;
    private javax.swing.JButton export;
    private javax.swing.JButton jButton1;
    private javax.swing.JButton jButton2;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JTextField name;
    private javax.swing.JTable table;
    private javax.swing.JTextField user_code;
    // End of variables declaration//GEN-END:variables
}
