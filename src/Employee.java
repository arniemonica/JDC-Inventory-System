
import java.awt.Color;
import java.awt.Desktop;
import java.awt.Font;
import java.awt.Image;
import java.awt.print.PrinterException;
import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.text.MessageFormat;
import java.text.SimpleDateFormat;
import java.util.Locale;
import java.util.Vector;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTabbedPane;
import javax.swing.JTable;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
/**
 *
 * @author Arnie D
 */
public class Employee extends javax.swing.JFrame {

    /**
     * Creates new form Employee
     */
    String inv = "false";

    public Employee() {
        initComponents();

        jLabel55.setVisible(false);
        jLabel63.setVisible(false);
        jLabel73.setVisible(false);

        Connect();

        Icon viber = jLabel90.getIcon();
        ImageIcon iconv = (ImageIcon) viber;
        Image imagev = iconv.getImage().getScaledInstance(jLabel90.getWidth(), jLabel90.getHeight(), Image.SCALE_SMOOTH);
        jLabel90.setIcon(new ImageIcon(imagev));

        Icon i = jLabel24.getIcon();
        ImageIcon icon = (ImageIcon) i;
        Image image = icon.getImage().getScaledInstance(jLabel24.getWidth(), jLabel24.getHeight(), Image.SCALE_SMOOTH);
        jLabel24.setIcon(new ImageIcon(image));

        Icon b = jLabel29.getIcon();
        ImageIcon iconq = (ImageIcon) b;
        Image imageq = iconq.getImage().getScaledInstance(jLabel29.getWidth(), jLabel29.getHeight(), Image.SCALE_SMOOTH);
        jLabel29.setIcon(new ImageIcon(imageq));

        Icon in = jLabel64.getIcon();
        ImageIcon iconin = (ImageIcon) in;
        Image imagein = iconin.getImage().getScaledInstance(jLabel64.getWidth(), jLabel64.getHeight(), Image.SCALE_SMOOTH);
        jLabel64.setIcon(new ImageIcon(imagein));

        Icon inl = jLabel65.getIcon();
        ImageIcon iconinl = (ImageIcon) inl;
        Image imageinl = iconinl.getImage().getScaledInstance(jLabel65.getWidth(), jLabel65.getHeight(), Image.SCALE_SMOOTH);
        jLabel65.setIcon(new ImageIcon(imagein));

    }

    Connection con;
    PreparedStatement pst;
    ResultSet rs;

    public void Connect() {
        try {
            Class.forName("com.mysql.cj.jdbc.Driver");
            con = DriverManager.getConnection("jdbc:mysql://localhost:3306/jdc", "root", "admin");
        } catch (ClassNotFoundException | SQLException ex) {
            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public void Import() {

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

                    DefaultTableModel df = (DefaultTableModel) user_table.getModel();
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

    public void Import_Inventory() {

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

                    XSSFCell eName = excelRow.getCell(0);
                    XSSFCell eDept = excelRow.getCell(1);
                    XSSFCell eMon_b = excelRow.getCell(2);
                    XSSFCell eMon_A = excelRow.getCell(3);
                    XSSFCell eMnK = excelRow.getCell(4);
                    XSSFCell eM_b = excelRow.getCell(5);
                    XSSFCell eM_S_N = excelRow.getCell(6);
                    XSSFCell ePSB = excelRow.getCell(7);
                    XSSFCell ePSSN = excelRow.getCell(8);
                    XSSFCell eHDB = excelRow.getCell(9);
                    XSSFCell eHDS = excelRow.getCell(10);
                    XSSFCell eHDS_N = excelRow.getCell(11);
                    XSSFCell eMB = excelRow.getCell(12);
                    XSSFCell eMS = excelRow.getCell(13);
                    XSSFCell eMSN = excelRow.getCell(14);
                    XSSFCell eGC = excelRow.getCell(15);
                    XSSFCell eSN = excelRow.getCell(16);
                    XSSFCell eP = excelRow.getCell(17);
                    XSSFCell ePS = excelRow.getCell(18);
                    XSSFCell eOLA = excelRow.getCell(19);
                    XSSFCell eWLA = excelRow.getCell(20);
                    XSSFCell eIP = excelRow.getCell(21);
                    XSSFCell eYT = excelRow.getCell(22);
                    XSSFCell eFB = excelRow.getCell(23);
                    XSSFCell eUSB = excelRow.getCell(24);
                    XSSFCell eD = excelRow.getCell(25);
                    XSSFCell eW = excelRow.getCell(26);
                    XSSFCell eH = excelRow.getCell(27);
                    XSSFCell eWEDP = excelRow.getCell(28);

                    DefaultTableModel df = (DefaultTableModel) computer_table.getModel();
                    df.addRow(new Object[]{eName, eDept, eMon_b, eMon_A, eMnK, eM_b, eM_S_N, ePSB, ePSSN, eHDB, eHDS, eHDS_N, eMB, eMS, eMSN, eGC, eSN, eP, ePS, eOLA, eWLA, eIP, eYT, eFB, eUSB, eD, eW, eH, eWEDP});

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

    public void Import_LaptopInventory() {

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
                    XSSFCell eAs_ID = excelRow.getCell(1);
                    XSSFCell eAs_Desc = excelRow.getCell(2);
                    XSSFCell eBrand = excelRow.getCell(3);
                    XSSFCell eModel = excelRow.getCell(4);
                    XSSFCell eS_Num = excelRow.getCell(5);
                    XSSFCell eAccount = excelRow.getCell(6);
                    XSSFCell eWar = excelRow.getCell(7);
                    XSSFCell eCond = excelRow.getCell(8);
                    XSSFCell eStats = excelRow.getCell(9);
                    XSSFCell eReco = excelRow.getCell(10);

                    DefaultTableModel df = (DefaultTableModel) laptop_table.getModel();
                    df.addRow(new Object[]{eDept, String.format("%.0f", eAs_ID.getNumericCellValue()),
                        eAs_Desc, eBrand, eModel, eS_Num, eAccount, eWar, eCond, eStats, eReco});

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

    public void Import_Viber() {

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
                    XSSFCell eCname = excelRow.getCell(0);
                    XSSFCell eMn = excelRow.getCell(1);
                    XSSFCell eDep = excelRow.getCell(2);
                    XSSFCell eDeT = excelRow.getCell(3);

                    DefaultTableModel df = (DefaultTableModel) viber_table.getModel();
                    df.addRow(new Object[]{eCname, eMn, eDep, eDeT});

                }
                int columnIndex = 1;
                DefaultTableModel model = (DefaultTableModel) viber_table.getModel();
                int rowCount = model.getRowCount();
                int saved = 0;
                int dup = 1;
                if (rowCount != 0) {
                    for (int row = 0; row < rowCount; row++) {
                        Object value = model.getValueAt(row, columnIndex);
                        try {

                            String duplicate = "SELECT * FROM viber_accounts WHERE mobile_number = '" + value + "'";
                            pst = con.prepareStatement(duplicate);
                            rs = pst.executeQuery(duplicate);

                            if (rs.next()) {
                                if (dup == 1) {
                                    dup = 0;
                                    JOptionPane.showMessageDialog(this, "Some Mobile number already used ");
                                    Fetch_Viber();

                                }

                            } else if (!rs.next()) {

                                int rowCount1 = viber_table.getRowCount();
                                int columnCount = viber_table.getColumnCount();
                                for (int row1 = 0; row1 < rowCount1; row1++) {

                                    Object val = viber_table.getValueAt(row1, 1);
                                    Object c_name = viber_table.getValueAt(row1, 0);
                                    Object num = viber_table.getValueAt(row1, 1);
                                    Object dept = viber_table.getValueAt(row1, 2);
                                    Object d_t = viber_table.getValueAt(row1, 3);

                                    if (c_name == null) {
                                        c_name = " ";
                                    }
                                    if (num == null) {
                                        num = " ";
                                    }
                                    if (dept == null) {
                                        dept = " ";
                                    }
                                    if (d_t == null) {
                                        d_t = " ";
                                    }

                                    if (val == value) {

                                        pst = con.prepareStatement("Insert into viber_accounts(client_name,mobile_number,department,device_type) values (?,?,?,?)");
                                        pst.setString(1, c_name.toString());
                                        pst.setString(2, num.toString());
                                        pst.setString(3, dept.toString());
                                        pst.setString(4, d_t.toString());

                                        saved = pst.executeUpdate();

                                    }

                                }

                            }

                        } catch (SQLException ex) {
                            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);
                        }

                    }
                    if (saved == 1) {
                        JOptionPane.showMessageDialog(null, "SAVED!");
                        Fetch_Viber();
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "You need to import files");
                }

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
 public void Import_Email() {

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
                    XSSFCell eName = excelRow.getCell(0);
                    XSSFCell ePos = excelRow.getCell(1);
                    XSSFCell eDep = excelRow.getCell(2);
                    XSSFCell eEm = excelRow.getCell(3);

                    DefaultTableModel df = (DefaultTableModel) email_table.getModel();
                    df.addRow(new Object[]{eName, ePos, eDep, eEm});

                }
                int columnIndex = 3;
                DefaultTableModel model = (DefaultTableModel) email_table.getModel();
                int rowCount = model.getRowCount();
                int saved = 0;
                int dup = 1;
                if (rowCount != 0) {
                    for (int row = 0; row < rowCount; row++) {
                        Object value = model.getValueAt(row, columnIndex);
                        try {

                            String duplicate = "SELECT * FROM email_list WHERE email = '" + value + "'";
                            pst = con.prepareStatement(duplicate);
                            rs = pst.executeQuery(duplicate);

                            if (rs.next()) {
                                if (dup == 1) {
                                    dup = 0;
                                    JOptionPane.showMessageDialog(this, "Email Account already used ");
                                    Fetch_Email();

                                }

                            } else if (!rs.next()) {

                                int rowCount1 = email_table.getRowCount();
                                int columnCount = email_table.getColumnCount();
                                for (int row1 = 0; row1 < rowCount1; row1++) {

                                    Object val = email_table.getValueAt(row1, 3);
                                    Object email_name = email_table.getValueAt(row1, 0);
                                    Object email_pos = email_table.getValueAt(row1, 1);
                                    Object email_dept = email_table.getValueAt(row1, 2);
                                    Object email_email = email_table.getValueAt(row1, 3);

                                    if (email_name == null) {
                                        email_name = " ";
                                    }
                                    if (email_pos == null) {
                                        email_pos = " ";
                                    }
                                    if (email_dept == null) {
                                        email_dept = " ";
                                    }
                                    if (email_email == null) {
                                        email_email = " ";
                                    }

                                    if (val == value) {

                                        pst = con.prepareStatement("Insert into email_list(Name,Position,department,email) values (?,?,?,?)");
                                        pst.setString(1, email_name.toString());
                                        pst.setString(2, email_pos.toString());
                                        pst.setString(3, email_dept.toString());
                                        pst.setString(4, email_email.toString());

                                        saved = pst.executeUpdate();

                                    }

                                }

                            }

                        } catch (SQLException ex) {
                            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);
                        }

                    }
                    if (saved == 1) {
                        JOptionPane.showMessageDialog(null, "SAVED!");
                        Fetch_Email();
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "You need to import files");
                }

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
    public void search_inv() {
        try {
            String s = search.getText();
            String a = Assets.getSelectedItem().toString();
            String search = "false";
            if (a.equals("Computer")) {
                pst = con.prepareStatement("Select * FROM Computer_Inventory WHERE NAME='" + s + "' or DEPARTMENT ='" + s + "' or MONITOR_BRAND='" + s
                        + "' or MONITOR_ASSET_BRAND='" + s + "' or MOUSE_AND_KEYBOARD='" + s + "' or MOTHERBOARD_BRAND_MODEL='" + s + "' or MOTHERBOARD_SERIAL_NO='" + s + "' or POWERSUPPLY_BRAND='" + s
                        + "' or POWERSUPPLY_SERIAL_NO='" + s + "' or HARD_DRIVE_BRAND='" + s + "' or HARD_DRIVE_SIZE='" + s + "'or HARD_DRIVE_SERIAL_NO ='" + s + "' or MEMORY_BRAND='" + s
                        + "' or MEMORY_SIZE='" + s + "' or MEMORY_SERIAL_NO='" + s + "' or GRAPHIC_CARDS='" + s + "' or SERIAL_NUMBER='" + s + "' or PROCESSOR='" + s
                        + "' or PROCESSOR_SPECS='" + s + "' or OFFICE_LICENSE_ACTIVATED='" + s + "' or WINDOWS_LICENSE_ACTIVATED='" + s + "'or IP_ADDRESS ='" + s + "' or YOUTUBE_BLOCKED='" + s
                        + "' or FB_BLOCKED='" + s + "' or USB_ENABLED='" + s + "' or DOMAIN='" + s + "' or WEBCAM='" + s + "' or HEADSET='" + s
                        + "' or WARRANTY_END_DATE_PROCESSOR='" + s + "' ");
                rs = pst.executeQuery();
                DefaultTableModel model = (DefaultTableModel) inv_search_table.getModel();
                model.setRowCount(0);

                while (rs.next()) {
                    search = "true";
                    String name = rs.getString("NAME");
                    String dept = rs.getString("DEPARTMENT");
                    String mbrand = rs.getString("MONITOR_BRAND");
                    String mab = rs.getString("MONITOR_ASSET_BRAND");
                    String mak = rs.getString("MOUSE_AND_KEYBOARD");
                    String mbm = rs.getString("MOTHERBOARD_BRAND_MODEL");
                    String msn = rs.getString("MOTHERBOARD_SERIAL_NO");
                    String pb = rs.getString("POWERSUPPLY_BRAND");
                    String psn = rs.getString("POWERSUPPLY_SERIAL_NO");
                    String hdb = rs.getString("HARD_DRIVE_BRAND");
                    String hds = rs.getString("HARD_DRIVE_SIZE");
                    String hdsn = rs.getString("HARD_DRIVE_SERIAL_NO");
                    String mb = rs.getString("MEMORY_BRAND");
                    String ms = rs.getString("MEMORY_SIZE");
                    String mmrysn = rs.getString("MEMORY_SERIAL_NO");
                    String gc = rs.getString("GRAPHIC_CARDS");
                    String sn = rs.getString("SERIAL_NUMBER");
                    String p = rs.getString("PROCESSOR");
                    String ps = rs.getString("PROCESSOR_SPECS");
                    String ola = rs.getString("OFFICE_LICENSE_ACTIVATED");
                    String wla = rs.getString("WINDOWS_LICENSE_ACTIVATED");
                    String ia = rs.getString("IP_ADDRESS");
                    String yb = rs.getString("YOUTUBE_BLOCKED");
                    String fb = rs.getString("FB_BLOCKED");
                    String usb = rs.getString("USB_ENABLED");
                    String d = rs.getString("DOMAIN");
                    String w = rs.getString("WEBCAM");
                    String h = rs.getString("HEADSET");
                    String wedp = rs.getString("WARRANTY_END_DATE_PROCESSOR");

                    model.addRow(new Object[]{name, dept, mbrand, mab, mak, mbm, msn, pb, psn, hdb, hds, hdsn, mb, ms, mmrysn, gc, sn, p, ps, ola, wla, ia, yb, fb, usb, d, w, h, wedp});

                }
            } else {
                pst = con.prepareStatement("Select * FROM Laptop_Inventory WHERE department='" + s + "' or asset_id ='" + s + "' or asset_description='" + s
                        + "' or brand='" + s + "' or model='" + s + "' or serial_number='" + s + "' or accountable_to='" + s + "' or warranty_date='" + s
                        + "' or conditions='" + s + "' or status='" + s + "' or recommendation='" + s + "'");
                rs = pst.executeQuery();
                DefaultTableModel model = (DefaultTableModel) inv_search_table.getModel();
                model.setRowCount(0);

                while (rs.next()) {
                    search = "true";
                    String dept = rs.getString("department");
                    String asst_id = rs.getString("asset_id");
                    String asst_desc = rs.getString("asset_description");
                    String brand = rs.getString("brand");
                    String models = rs.getString("model");
                    String s_num = rs.getString("serial_number");
                    String acc_to = rs.getString("accountable_to");
                    String warranty = rs.getString("warranty_date");
                    String cond = rs.getString("conditions");
                    String stats = rs.getString("status");
                    String reco = rs.getString("recommendation");

                    model.addRow(new Object[]{dept, asst_id, asst_desc, brand, models, s_num, acc_to, warranty, cond, stats, reco});

                }
            }

            if (search == "false" && !rs.next()) {
                JOptionPane.showMessageDialog(null, "No user code found");
                Fetch1();
            }

        } catch (Exception e) {
            System.out.print(e);

        }
    }

    public void search_viber() {
        try {
            String v = viber_search.getText();
            String search = "false";

            pst = con.prepareStatement("Select * FROM viber_accounts WHERE client_name='" + v + "' or mobile_number ='" + v + "' or department='" + v
                    + "' or device_type='" + v + "'");
            rs = pst.executeQuery();
            DefaultTableModel model = (DefaultTableModel) viber_table.getModel();
            model.setRowCount(0);

            while (rs.next()) {
                search = "true";
                String name = rs.getString("client_name");
                String num = rs.getString("mobile_number");
                String dep = rs.getString("department");
                String dt = rs.getString("device_type");

                model.addRow(new Object[]{name, num, dep, dt});

            }

            if (search == "false" && !rs.next()) {
                JOptionPane.showMessageDialog(null, "No search result found");
                Fetch1();
            }

        } catch (Exception e) {
            System.out.print(e);

        }
    }
    public void search_email() {
        try {
            String v = email_search.getText();
            String search = "false";

            pst = con.prepareStatement("Select * FROM email_list WHERE name='" + v + "' or position='" + v + "' or department='" + v
                    + "' or email='" + v + "'");
            rs = pst.executeQuery();
            DefaultTableModel model = (DefaultTableModel) email_table.getModel();
            model.setRowCount(0);

            while (rs.next()) {
                search = "true";
                String name = rs.getString("name");
                String pos = rs.getString("position");
                String dep = rs.getString("department");
                String email = rs.getString("email");

                model.addRow(new Object[]{name, pos, dep, email});

            }

            if (search == "false" && !rs.next()) {
                JOptionPane.showMessageDialog(null, "No search result found");
                Fetch_Email();
            }

        } catch (Exception e) {
            System.out.print(e);

        }
    }

    public void search_user() {
        try {
            String v = user_search.getText();
            String search = "false";

            pst = con.prepareStatement("Select * FROM user WHERE department='" + v + "' or activity ='" + v + "' or db='" + v
                    + "' or user_code='" + v + "' or name='" + v + "' or license ='" + v + "' or remarks='" + v + "'");
            rs = pst.executeQuery();
            DefaultTableModel model = (DefaultTableModel) search_table.getModel();
            model.setRowCount(0);

            while (rs.next()) {
                search = "true";
                String dept = rs.getString("department");
                String act = rs.getString("activity");
                String db = rs.getString("db");
                String user = rs.getString("user_code");
                String nm = rs.getString("name");
                String l = rs.getString("license");
                String r = rs.getString("remarks");

                model.addRow(new Object[]{dept, act, db, user, nm, l, r});

            }

            if (search == "false" && !rs.next()) {
                JOptionPane.showMessageDialog(null, "No search result found");
                isEmpty();
            }

        } catch (Exception e) {
            System.out.print(e);

        }
    }

    public void remove() {
        user_code.setText("");
        user_name.setText("");
        user_department.setText("");
        user_activity.setText("");
        user_database.setText("");
        LType.setSelectedIndex(0);
        user_remark.setText("");
    }

    public void isEmpty() {
        try {
            pst = con.prepareStatement("Select * from user");
            rs = pst.executeQuery();
            DefaultTableModel model = (DefaultTableModel) search_table.getModel();
            Object[] row;
            while (rs.next()) {
                row = new Object[7];
                row[0] = rs.getString(2);
                row[1] = rs.getString(3);
                row[2] = rs.getString(4);
                row[3] = rs.getString(5);
                row[4] = rs.getString(6);
                row[5] = rs.getString(7);
                row[6] = rs.getString(8);

                model.addRow(row);
            }
            
        } catch (Exception ex) {
            ex.printStackTrace();   
        } finally {
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

    public void openFile(String file) {
        try {
            File path = new File(file);
            Desktop.getDesktop().open(path);
        } catch (IOException ioe) {
            System.out.println(ioe);
        }
    }

    private void Fetch() {
        try {
            int a;
            pst = con.prepareStatement("SELECT * FROM user");
            rs = pst.executeQuery();
            ResultSetMetaData rss = rs.getMetaData();
            a = rss.getColumnCount();

            DefaultTableModel df = (DefaultTableModel) search_table.getModel();
            df.setRowCount(0);
            while (rs.next()) {
                Vector v2 = new Vector();
                for (int x = 1; x <= a; x++) {
                    v2.add(rs.getString("department"));
                    v2.add(rs.getString("activity"));
                    v2.add(rs.getString("db"));
                    v2.add(rs.getString("user_code"));
                    v2.add(rs.getString("name"));
                    v2.add(rs.getString("license"));
                    v2.add(rs.getString("remarks"));
                }
                df.addRow(v2);
            }

        } catch (SQLException ex) {
            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    private void Fetch_Viber() {
        try {
            int a;
            pst = con.prepareStatement("SELECT * FROM viber_accounts");
            rs = pst.executeQuery();
            ResultSetMetaData rss = rs.getMetaData();
            a = rss.getColumnCount();

            DefaultTableModel df = (DefaultTableModel) viber_table.getModel();
            df.setRowCount(0);
            while (rs.next()) {
                Vector v2 = new Vector();
                for (int x = 1; x <= a; x++) {
                    v2.add(rs.getString("client_name"));
                    v2.add(rs.getString("mobile_number"));
                    v2.add(rs.getString("department"));
                    v2.add(rs.getString("device_type"));
                }
                df.addRow(v2);
                Viber_add.setEnabled(false);
                Viber_edit.setEnabled(false);
                Delete.setEnabled(false);
                client_name.setText("");
                mobile_number.setText("");
                department.setText("");
            }

        } catch (SQLException ex) {
            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
       private void Fetch_Email() {
        try {
            int a;
            pst = con.prepareStatement("SELECT * FROM email_list");
            rs = pst.executeQuery();
            ResultSetMetaData rss = rs.getMetaData();
            a = rss.getColumnCount();

            DefaultTableModel df = (DefaultTableModel) email_table.getModel();
            df.setRowCount(0);
            while (rs.next()) {
                Vector v2 = new Vector();
                for (int x = 1; x <= a; x++) {
                    v2.add(rs.getString("name"));
                    v2.add(rs.getString("position"));
                    v2.add(rs.getString("department"));
                    v2.add(rs.getString("email"));
                }
                df.addRow(v2);
                email_add.setEnabled(false);
                email_edit.setEnabled(false);
                email_delete.setEnabled(false);
                email_name.setText("");
                email_position.setText("");
                email_email.setText("");
                email_department.setSelectedIndex(0);
            }

        } catch (SQLException ex) {
            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public void Fetch1() {

        try {
            String s = Assets.getSelectedItem().toString();
            int a;
            if (s.equals("Computer")) {
                pst = con.prepareStatement("SELECT * FROM computer_inventory");
                rs = pst.executeQuery();
                ResultSetMetaData rss = rs.getMetaData();
                a = rss.getColumnCount();

                DefaultTableModel df = (DefaultTableModel) inv_search_table.getModel();
                df.setRowCount(0);
                while (rs.next()) {
                    Vector v2 = new Vector();
                    for (int x = 1; x <= a; x++) {
                        v2.add(rs.getString("NAME"));
                        v2.add(rs.getString("DEPARTMENT"));
                        v2.add(rs.getString("MONITOR_BRAND"));
                        v2.add(rs.getString("MONITOR_ASSET_BRAND"));
                        v2.add(rs.getString("MOUSE_AND_KEYBOARD"));
                        v2.add(rs.getString("MOTHERBOARD_BRAND_MODEL"));
                        v2.add(rs.getString("MOTHERBOARD_SERIAL_NO"));
                        v2.add(rs.getString("POWERSUPPLY_BRAND"));
                        v2.add(rs.getString("POWERSUPPLY_SERIAL_NO"));
                        v2.add(rs.getString("HARD_DRIVE_BRAND"));
                        v2.add(rs.getString("HARD_DRIVE_SIZE"));
                        v2.add(rs.getString("HARD_DRIVE_SERIAL_NO"));
                        v2.add(rs.getString("MEMORY_BRAND"));
                        v2.add(rs.getString("MEMORY_SIZE"));
                        v2.add(rs.getString("MEMORY_SERIAL_NO"));
                        v2.add(rs.getString("GRAPHIC_CARDS"));
                        v2.add(rs.getString("SERIAL_NUMBER"));
                        v2.add(rs.getString("PROCESSOR"));
                        v2.add(rs.getString("PROCESSOR_SPECS"));
                        v2.add(rs.getString("OFFICE_LICENSE_ACTIVATED"));
                        v2.add(rs.getString("WINDOWS_LICENSE_ACTIVATED"));
                        v2.add(rs.getString("IP_ADDRESS"));
                        v2.add(rs.getString("YOUTUBE_BLOCKED"));
                        v2.add(rs.getString("FB_BLOCKED"));
                        v2.add(rs.getString("USB_ENABLED"));
                        v2.add(rs.getString("DOMAIN"));
                        v2.add(rs.getString("WEBCAM"));
                        v2.add(rs.getString("HEADSET"));
                        v2.add(rs.getString("WARRANTY_END_DATE_PROCESSOR"));

                    }
                    df.addRow(v2);
                }
            } else {
                pst = con.prepareStatement("SELECT * FROM laptop_inventory");
                rs = pst.executeQuery();
                ResultSetMetaData rss = rs.getMetaData();
                a = rss.getColumnCount();

                DefaultTableModel df = (DefaultTableModel) inv_search_table.getModel();
                df.setRowCount(0);
                while (rs.next()) {
                    Vector v2 = new Vector();
                    for (int x = 1; x <= a; x++) {
                        v2.add(rs.getString("department"));
                        v2.add(rs.getString("asset_id"));
                        v2.add(rs.getString("asset_description"));
                        v2.add(rs.getString("brand"));
                        v2.add(rs.getString("model"));
                        v2.add(rs.getString("serial_number"));
                        v2.add(rs.getString("accountable_to"));
                        v2.add(rs.getString("warranty_date"));
                        v2.add(rs.getString("conditions"));
                        v2.add(rs.getString("status"));
                        v2.add(rs.getString("recommendation"));

                    }
                    df.addRow(v2);
                }
            }

        } catch (SQLException ex) {
            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);
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
        home_tab = new javax.swing.JPanel();
        jLabel1 = new javax.swing.JLabel();
        jLabel2 = new javax.swing.JLabel();
        home_p1 = new javax.swing.JPanel();
        jLabel32 = new javax.swing.JLabel();
        jLabel33 = new javax.swing.JLabel();
        user_tab = new javax.swing.JPanel();
        jLabel14 = new javax.swing.JLabel();
        jLabel15 = new javax.swing.JLabel();
        search_tab = new javax.swing.JPanel();
        jLabel30 = new javax.swing.JLabel();
        jPanel9 = new javax.swing.JPanel();
        jLabel31 = new javax.swing.JLabel();
        jLabel34 = new javax.swing.JLabel();
        jPanel12 = new javax.swing.JPanel();
        jLabel35 = new javax.swing.JLabel();
        jLabel36 = new javax.swing.JLabel();
        jPanel13 = new javax.swing.JPanel();
        jLabel37 = new javax.swing.JLabel();
        jLabel38 = new javax.swing.JLabel();
        jLabel39 = new javax.swing.JLabel();
        inventory1 = new javax.swing.JPanel();
        jLabel40 = new javax.swing.JLabel();
        jPanel15 = new javax.swing.JPanel();
        jLabel41 = new javax.swing.JLabel();
        jLabel42 = new javax.swing.JLabel();
        jPanel16 = new javax.swing.JPanel();
        jLabel43 = new javax.swing.JLabel();
        jLabel44 = new javax.swing.JLabel();
        jPanel17 = new javax.swing.JPanel();
        jLabel45 = new javax.swing.JLabel();
        jLabel46 = new javax.swing.JLabel();
        jLabel47 = new javax.swing.JLabel();
        lap_tab = new javax.swing.JPanel();
        jLabel48 = new javax.swing.JLabel();
        jPanel18 = new javax.swing.JPanel();
        jLabel49 = new javax.swing.JLabel();
        jLabel50 = new javax.swing.JLabel();
        jPanel19 = new javax.swing.JPanel();
        jLabel51 = new javax.swing.JLabel();
        jLabel52 = new javax.swing.JLabel();
        jPanel20 = new javax.swing.JPanel();
        jLabel53 = new javax.swing.JLabel();
        jLabel54 = new javax.swing.JLabel();
        jLabel63 = new javax.swing.JLabel();
        com_tab = new javax.swing.JPanel();
        jPanel21 = new javax.swing.JPanel();
        jLabel56 = new javax.swing.JLabel();
        jLabel57 = new javax.swing.JLabel();
        jPanel22 = new javax.swing.JPanel();
        jLabel58 = new javax.swing.JLabel();
        jLabel59 = new javax.swing.JLabel();
        jPanel23 = new javax.swing.JPanel();
        jLabel60 = new javax.swing.JLabel();
        jLabel61 = new javax.swing.JLabel();
        jLabel62 = new javax.swing.JLabel();
        jLabel55 = new javax.swing.JLabel();
        jPanel5 = new javax.swing.JPanel();
        jLabel6 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        inv_search_tab = new javax.swing.JPanel();
        jLabel66 = new javax.swing.JLabel();
        jPanel24 = new javax.swing.JPanel();
        jLabel67 = new javax.swing.JLabel();
        jLabel68 = new javax.swing.JLabel();
        jPanel25 = new javax.swing.JPanel();
        jLabel69 = new javax.swing.JLabel();
        jLabel70 = new javax.swing.JLabel();
        jPanel26 = new javax.swing.JPanel();
        jLabel71 = new javax.swing.JLabel();
        jLabel72 = new javax.swing.JLabel();
        jLabel73 = new javax.swing.JLabel();
        viber_tab = new javax.swing.JPanel();
        jLabel89 = new javax.swing.JLabel();
        jLabel90 = new javax.swing.JLabel();
        email_tab = new javax.swing.JPanel();
        jLabel96 = new javax.swing.JLabel();
        jLabel97 = new javax.swing.JLabel();
        jPanel14 = new javax.swing.JPanel();
        jLabel28 = new javax.swing.JLabel();
        jLabel120 = new javax.swing.JLabel();
        jLabel121 = new javax.swing.JLabel();
        jLabel20 = new javax.swing.JLabel();
        jTabbedPane1 = new javax.swing.JTabbedPane();
        jPanel4 = new javax.swing.JPanel();
        jPanel2 = new javax.swing.JPanel();
        jLabel18 = new javax.swing.JLabel();
        jLabel23 = new javax.swing.JLabel();
        jLabel24 = new javax.swing.JLabel();
        jLabel25 = new javax.swing.JLabel();
        jLabel26 = new javax.swing.JLabel();
        jLabel27 = new javax.swing.JLabel();
        jPanel10 = new javax.swing.JPanel();
        jScrollPane1 = new javax.swing.JScrollPane();
        search_table = new javax.swing.JTable();
        jLabel5 = new javax.swing.JLabel();
        user_code = new javax.swing.JTextField();
        jLabel16 = new javax.swing.JLabel();
        user_name = new javax.swing.JTextField();
        user_department = new javax.swing.JTextField();
        jLabel17 = new javax.swing.JLabel();
        LType = new javax.swing.JComboBox<>();
        SaveDB2 = new javax.swing.JButton();
        user_remove = new javax.swing.JButton();
        jLabel19 = new javax.swing.JLabel();
        user_activity = new javax.swing.JTextField();
        jLabel21 = new javax.swing.JLabel();
        user_database = new javax.swing.JTextField();
        jLabel22 = new javax.swing.JLabel();
        user_remark = new javax.swing.JTextField();
        l_print2 = new javax.swing.JButton();
        user_edit = new javax.swing.JButton();
        jPanel39 = new javax.swing.JPanel();
        user_search = new javax.swing.JTextField();
        SaveDB21 = new javax.swing.JButton();
        jLabel95 = new javax.swing.JLabel();
        jPanel11 = new javax.swing.JPanel();
        jScrollPane2 = new javax.swing.JScrollPane();
        user_table = new javax.swing.JTable();
        jButton1 = new javax.swing.JButton();
        export = new javax.swing.JButton();
        SaveDB = new javax.swing.JButton();
        jLabel29 = new javax.swing.JLabel();
        SaveDB6 = new javax.swing.JButton();
        jTabbedPane2 = new javax.swing.JTabbedPane();
        COMPUTER = new javax.swing.JPanel();
        jLabel64 = new javax.swing.JLabel();
        c_import = new javax.swing.JButton();
        c_export = new javax.swing.JButton();
        c_save = new javax.swing.JButton();
        jScrollPane3 = new javax.swing.JScrollPane();
        computer_table = new javax.swing.JTable();
        c_print = new javax.swing.JButton();
        jPanel3 = new javax.swing.JPanel();
        jScrollPane5 = new javax.swing.JScrollPane();
        laptop_table = new javax.swing.JTable();
        jLabel65 = new javax.swing.JLabel();
        l_import = new javax.swing.JButton();
        l_export = new javax.swing.JButton();
        l_save = new javax.swing.JButton();
        l_print = new javax.swing.JButton();
        jPanel6 = new javax.swing.JPanel();
        Assets = new javax.swing.JComboBox<>();
        SaveDB10 = new javax.swing.JButton();
        search = new javax.swing.JTextField();
        jScrollPane6 = new javax.swing.JScrollPane();
        inv_search_table = new javax.swing.JTable();
        SaveDB9 = new javax.swing.JButton();
        SaveDB11 = new javax.swing.JButton();
        l_print1 = new javax.swing.JButton();
        jPanel30 = new javax.swing.JPanel();
        jPanel31 = new javax.swing.JPanel();
        jPanel32 = new javax.swing.JPanel();
        jPanel33 = new javax.swing.JPanel();
        jPanel34 = new javax.swing.JPanel();
        c_n = new javax.swing.JTextField();
        c_dept = new javax.swing.JTextField();
        c_mak = new javax.swing.JTextField();
        c_mn = new javax.swing.JTextField();
        c_mab = new javax.swing.JTextField();
        c_hds = new javax.swing.JTextField();
        c_pssn = new javax.swing.JTextField();
        c_psb = new javax.swing.JTextField();
        c_mbm = new javax.swing.JTextField();
        c_hdb = new javax.swing.JTextField();
        SaveDB13 = new javax.swing.JButton();
        jLabel78 = new javax.swing.JLabel();
        jLabel79 = new javax.swing.JLabel();
        jLabel80 = new javax.swing.JLabel();
        jLabel81 = new javax.swing.JLabel();
        jLabel82 = new javax.swing.JLabel();
        jLabel83 = new javax.swing.JLabel();
        jLabel84 = new javax.swing.JLabel();
        jLabel85 = new javax.swing.JLabel();
        jLabel86 = new javax.swing.JLabel();
        jLabel87 = new javax.swing.JLabel();
        jLabel88 = new javax.swing.JLabel();
        c_ms = new javax.swing.JTextField();
        jLabel100 = new javax.swing.JLabel();
        jLabel101 = new javax.swing.JLabel();
        c_mb = new javax.swing.JTextField();
        c_hdsn = new javax.swing.JTextField();
        jLabel102 = new javax.swing.JLabel();
        jLabel103 = new javax.swing.JLabel();
        c_mmrysn = new javax.swing.JTextField();
        jLabel104 = new javax.swing.JLabel();
        c_gc = new javax.swing.JTextField();
        jLabel105 = new javax.swing.JLabel();
        c_sn = new javax.swing.JTextField();
        jLabel106 = new javax.swing.JLabel();
        c_p = new javax.swing.JTextField();
        jLabel107 = new javax.swing.JLabel();
        c_ps = new javax.swing.JTextField();
        jLabel108 = new javax.swing.JLabel();
        jLabel109 = new javax.swing.JLabel();
        jLabel110 = new javax.swing.JLabel();
        c_ip = new javax.swing.JTextField();
        jLabel111 = new javax.swing.JLabel();
        jLabel112 = new javax.swing.JLabel();
        jLabel113 = new javax.swing.JLabel();
        jLabel114 = new javax.swing.JLabel();
        jLabel115 = new javax.swing.JLabel();
        jLabel116 = new javax.swing.JLabel();
        jLabel117 = new javax.swing.JLabel();
        c_msn = new javax.swing.JTextField();
        jPanel35 = new javax.swing.JPanel();
        jPanel36 = new javax.swing.JPanel();
        c_wla = new javax.swing.JComboBox<>();
        c_oal = new javax.swing.JComboBox<>();
        c_yt = new javax.swing.JComboBox<>();
        c_fb = new javax.swing.JComboBox<>();
        c_usb = new javax.swing.JComboBox<>();
        c_d = new javax.swing.JComboBox<>();
        c_h = new javax.swing.JComboBox<>();
        c_w = new javax.swing.JComboBox<>();
        c_wedp = new com.toedter.calendar.JDateChooser();
        jPanel7 = new javax.swing.JPanel();
        jPanel8 = new javax.swing.JPanel();
        jPanel27 = new javax.swing.JPanel();
        jPanel28 = new javax.swing.JPanel();
        jPanel29 = new javax.swing.JPanel();
        dept1 = new javax.swing.JTextField();
        id = new javax.swing.JTextField();
        mdl = new javax.swing.JTextField();
        desc = new javax.swing.JTextField();
        brnd = new javax.swing.JTextField();
        reco = new javax.swing.JTextField();
        condi = new javax.swing.JTextField();
        acct = new javax.swing.JTextField();
        srl = new javax.swing.JTextField();
        status = new javax.swing.JTextField();
        SaveDB12 = new javax.swing.JButton();
        jLabel7 = new javax.swing.JLabel();
        jLabel8 = new javax.swing.JLabel();
        jLabel9 = new javax.swing.JLabel();
        jLabel10 = new javax.swing.JLabel();
        jLabel11 = new javax.swing.JLabel();
        jLabel12 = new javax.swing.JLabel();
        jLabel13 = new javax.swing.JLabel();
        jLabel74 = new javax.swing.JLabel();
        jLabel75 = new javax.swing.JLabel();
        jLabel76 = new javax.swing.JLabel();
        jLabel77 = new javax.swing.JLabel();
        date = new com.toedter.calendar.JDateChooser();
        jPanel37 = new javax.swing.JPanel();
        jLabel91 = new javax.swing.JLabel();
        client_name = new javax.swing.JTextField();
        jLabel92 = new javax.swing.JLabel();
        mobile_number = new javax.swing.JTextField();
        jLabel93 = new javax.swing.JLabel();
        department = new javax.swing.JTextField();
        SaveDB14 = new javax.swing.JButton();
        Viber_add = new javax.swing.JButton();
        Export_Viber = new javax.swing.JButton();
        Delete = new javax.swing.JButton();
        Viber_edit = new javax.swing.JButton();
        jScrollPane7 = new javax.swing.JScrollPane();
        viber_table = new javax.swing.JTable();
        Viber_Import = new javax.swing.JButton();
        viber_search = new javax.swing.JTextField();
        SaveDB20 = new javax.swing.JButton();
        jPanel38 = new javax.swing.JPanel();
        device_type = new javax.swing.JComboBox<>();
        jLabel94 = new javax.swing.JLabel();
        jPanel40 = new javax.swing.JPanel();
        jLabel98 = new javax.swing.JLabel();
        email_name = new javax.swing.JTextField();
        jLabel99 = new javax.swing.JLabel();
        email_position = new javax.swing.JTextField();
        jLabel118 = new javax.swing.JLabel();
        email_email = new javax.swing.JTextField();
        jLabel119 = new javax.swing.JLabel();
        email_department = new javax.swing.JComboBox<>();
        email_add = new javax.swing.JButton();
        email_import = new javax.swing.JButton();
        email_export = new javax.swing.JButton();
        email_delete = new javax.swing.JButton();
        email_edit = new javax.swing.JButton();
        email_print = new javax.swing.JButton();
        jPanel41 = new javax.swing.JPanel();
        email_search = new javax.swing.JTextField();
        email_searchbutton = new javax.swing.JButton();
        jScrollPane8 = new javax.swing.JScrollPane();
        email_table = new javax.swing.JTable();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setResizable(false);

        jPanel1.setBackground(new java.awt.Color(228, 57, 39));
        jPanel1.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        home_tab.setBackground(new java.awt.Color(195, 0, 0));
        home_tab.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                home_tabMouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                home_tabMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                home_tabMouseExited(evt);
            }
        });
        home_tab.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel1.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel1.setForeground(new java.awt.Color(255, 255, 255));
        jLabel1.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel1.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Home_2.png"))); // NOI18N
        home_tab.add(jLabel1, new org.netbeans.lib.awtextra.AbsoluteConstraints(30, 0, 60, 60));

        jLabel2.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel2.setForeground(new java.awt.Color(255, 255, 255));
        jLabel2.setText("HOME");
        home_tab.add(jLabel2, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        home_p1.setBackground(new java.awt.Color(195, 0, 0));
        home_p1.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                home_p1MouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                home_p1MouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                home_p1MouseExited(evt);
            }
        });
        home_p1.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel32.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel32.setForeground(new java.awt.Color(255, 255, 255));
        jLabel32.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel32.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Home_1.png"))); // NOI18N
        home_p1.add(jLabel32, new org.netbeans.lib.awtextra.AbsoluteConstraints(30, 0, 60, 60));

        jLabel33.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel33.setForeground(new java.awt.Color(255, 255, 255));
        jLabel33.setText("HOME");
        home_p1.add(jLabel33, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        home_tab.add(home_p1, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 110, 300, 60));

        jPanel1.add(home_tab, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 10, 310, 60));

        user_tab.setBackground(new java.awt.Color(228, 57, 39));
        user_tab.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                user_tabMouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                user_tabMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                user_tabMouseExited(evt);
            }
            public void mousePressed(java.awt.event.MouseEvent evt) {
                user_tabMousePressed(evt);
            }
        });
        user_tab.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel14.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel14.setForeground(new java.awt.Color(255, 255, 255));
        jLabel14.setText("SAP USER");
        jLabel14.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jLabel14MouseClicked(evt);
            }
        });
        user_tab.add(jLabel14, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        jLabel15.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel15.setForeground(new java.awt.Color(255, 255, 255));
        jLabel15.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel15.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Database_2.png"))); // NOI18N
        jLabel15.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jLabel15MouseClicked(evt);
            }
        });
        user_tab.add(jLabel15, new org.netbeans.lib.awtextra.AbsoluteConstraints(30, 0, 60, 60));

        jPanel1.add(user_tab, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 130, 310, 60));

        search_tab.setBackground(new java.awt.Color(228, 57, 39));
        search_tab.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                search_tabMouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                search_tabMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                search_tabMouseExited(evt);
            }
            public void mousePressed(java.awt.event.MouseEvent evt) {
                search_tabMousePressed(evt);
            }
        });
        search_tab.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel30.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel30.setForeground(new java.awt.Color(255, 255, 255));
        jLabel30.setText("SEARCH");
        jLabel30.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jLabel30MouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                jLabel30MouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                jLabel30MouseExited(evt);
            }
        });
        search_tab.add(jLabel30, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        jPanel9.setBackground(new java.awt.Color(228, 57, 39));
        jPanel9.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel31.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel31.setForeground(new java.awt.Color(255, 255, 255));
        jLabel31.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel31.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Home_1.png"))); // NOI18N
        jPanel9.add(jLabel31, new org.netbeans.lib.awtextra.AbsoluteConstraints(40, 10, 50, 40));

        jLabel34.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel34.setForeground(new java.awt.Color(255, 255, 255));
        jLabel34.setText("HOME");
        jPanel9.add(jLabel34, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        search_tab.add(jPanel9, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 60, 300, 60));

        jPanel12.setBackground(new java.awt.Color(228, 57, 39));
        jPanel12.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel35.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel35.setForeground(new java.awt.Color(255, 255, 255));
        jLabel35.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel35.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Home_1.png"))); // NOI18N
        jPanel12.add(jLabel35, new org.netbeans.lib.awtextra.AbsoluteConstraints(40, 10, 50, 40));

        jLabel36.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel36.setForeground(new java.awt.Color(255, 255, 255));
        jLabel36.setText("HOME");
        jPanel12.add(jLabel36, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        jPanel13.setBackground(new java.awt.Color(228, 57, 39));
        jPanel13.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel37.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel37.setForeground(new java.awt.Color(255, 255, 255));
        jLabel37.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel37.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Home_1.png"))); // NOI18N
        jPanel13.add(jLabel37, new org.netbeans.lib.awtextra.AbsoluteConstraints(40, 10, 50, 40));

        jLabel38.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel38.setForeground(new java.awt.Color(255, 255, 255));
        jLabel38.setText("HOME");
        jPanel13.add(jLabel38, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        jPanel12.add(jPanel13, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 60, 300, 60));

        search_tab.add(jPanel12, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 60, 300, 60));

        jLabel39.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel39.setForeground(new java.awt.Color(255, 255, 255));
        jLabel39.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel39.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Search_7.png"))); // NOI18N
        jLabel39.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jLabel39MouseClicked(evt);
            }
        });
        search_tab.add(jLabel39, new org.netbeans.lib.awtextra.AbsoluteConstraints(30, 0, 60, 60));

        jPanel1.add(search_tab, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 70, 310, 60));

        inventory1.setBackground(new java.awt.Color(228, 57, 39));
        inventory1.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                inventory1MouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                inventory1MouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                inventory1MouseExited(evt);
            }
            public void mousePressed(java.awt.event.MouseEvent evt) {
                inventory1MousePressed(evt);
            }
        });
        inventory1.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel40.setFont(new java.awt.Font("Segoe UI", 0, 16)); // NOI18N
        jLabel40.setForeground(new java.awt.Color(255, 255, 255));
        jLabel40.setText("IT ASSETS INVENTORY");
        jLabel40.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jLabel40MouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                jLabel40MouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                jLabel40MouseExited(evt);
            }
        });
        inventory1.add(jLabel40, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        jPanel15.setBackground(new java.awt.Color(228, 57, 39));
        jPanel15.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel41.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel41.setForeground(new java.awt.Color(255, 255, 255));
        jLabel41.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel41.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Home_1.png"))); // NOI18N
        jPanel15.add(jLabel41, new org.netbeans.lib.awtextra.AbsoluteConstraints(40, 10, 50, 40));

        jLabel42.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel42.setForeground(new java.awt.Color(255, 255, 255));
        jLabel42.setText("HOME");
        jPanel15.add(jLabel42, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        inventory1.add(jPanel15, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 60, 300, 60));

        jPanel16.setBackground(new java.awt.Color(228, 57, 39));
        jPanel16.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel43.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel43.setForeground(new java.awt.Color(255, 255, 255));
        jLabel43.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel43.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Home_1.png"))); // NOI18N
        jPanel16.add(jLabel43, new org.netbeans.lib.awtextra.AbsoluteConstraints(40, 10, 50, 40));

        jLabel44.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel44.setForeground(new java.awt.Color(255, 255, 255));
        jLabel44.setText("HOME");
        jPanel16.add(jLabel44, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        jPanel17.setBackground(new java.awt.Color(228, 57, 39));
        jPanel17.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel45.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel45.setForeground(new java.awt.Color(255, 255, 255));
        jLabel45.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel45.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Home_1.png"))); // NOI18N
        jPanel17.add(jLabel45, new org.netbeans.lib.awtextra.AbsoluteConstraints(40, 10, 50, 40));

        jLabel46.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel46.setForeground(new java.awt.Color(255, 255, 255));
        jLabel46.setText("HOME");
        jPanel17.add(jLabel46, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        jPanel16.add(jPanel17, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 60, 300, 60));

        inventory1.add(jPanel16, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 60, 300, 60));

        jLabel47.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel47.setForeground(new java.awt.Color(255, 255, 255));
        jLabel47.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel47.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Products Pile_1.png"))); // NOI18N
        jLabel47.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jLabel47MouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                jLabel47MouseEntered(evt);
            }
        });
        inventory1.add(jLabel47, new org.netbeans.lib.awtextra.AbsoluteConstraints(30, 0, 60, 60));

        jPanel1.add(inventory1, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 340, 310, 60));

        lap_tab.setBackground(new java.awt.Color(228, 57, 39));
        lap_tab.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                lap_tabMouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                lap_tabMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                lap_tabMouseExited(evt);
            }
            public void mousePressed(java.awt.event.MouseEvent evt) {
                lap_tabMousePressed(evt);
            }
        });
        lap_tab.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel48.setFont(new java.awt.Font("Segoe UI", 0, 16)); // NOI18N
        jLabel48.setForeground(new java.awt.Color(228, 57, 39));
        jLabel48.setText("LAPTOP");
        jLabel48.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jLabel48MouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                jLabel48MouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                jLabel48MouseExited(evt);
            }
        });
        lap_tab.add(jLabel48, new org.netbeans.lib.awtextra.AbsoluteConstraints(186, 5, -1, -1));

        jPanel18.setBackground(new java.awt.Color(228, 57, 39));
        jPanel18.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel49.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel49.setForeground(new java.awt.Color(255, 255, 255));
        jLabel49.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel49.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Home_1.png"))); // NOI18N
        jPanel18.add(jLabel49, new org.netbeans.lib.awtextra.AbsoluteConstraints(40, 10, 50, 40));

        jLabel50.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel50.setForeground(new java.awt.Color(255, 255, 255));
        jLabel50.setText("HOME");
        jPanel18.add(jLabel50, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        lap_tab.add(jPanel18, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 60, 300, 60));

        jPanel19.setBackground(new java.awt.Color(228, 57, 39));
        jPanel19.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel51.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel51.setForeground(new java.awt.Color(255, 255, 255));
        jLabel51.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel51.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Home_1.png"))); // NOI18N
        jPanel19.add(jLabel51, new org.netbeans.lib.awtextra.AbsoluteConstraints(40, 10, 50, 40));

        jLabel52.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel52.setForeground(new java.awt.Color(255, 255, 255));
        jLabel52.setText("HOME");
        jPanel19.add(jLabel52, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        jPanel20.setBackground(new java.awt.Color(228, 57, 39));
        jPanel20.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel53.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel53.setForeground(new java.awt.Color(255, 255, 255));
        jLabel53.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel53.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Home_1.png"))); // NOI18N
        jPanel20.add(jLabel53, new org.netbeans.lib.awtextra.AbsoluteConstraints(40, 10, 50, 40));

        jLabel54.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel54.setForeground(new java.awt.Color(255, 255, 255));
        jLabel54.setText("HOME");
        jPanel20.add(jLabel54, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        jPanel19.add(jPanel20, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 60, 300, 60));

        lap_tab.add(jPanel19, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 60, 300, 60));

        jLabel63.setFont(new java.awt.Font("Segoe UI", 0, 16)); // NOI18N
        jLabel63.setForeground(new java.awt.Color(228, 57, 39));
        jLabel63.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Laptop_2.png"))); // NOI18N
        jLabel63.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jLabel63MouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                jLabel63MouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                jLabel63MouseExited(evt);
            }
        });
        lap_tab.add(jLabel63, new org.netbeans.lib.awtextra.AbsoluteConstraints(130, 0, 50, 30));

        jPanel1.add(lap_tab, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 440, 310, 30));

        com_tab.setBackground(new java.awt.Color(228, 57, 39));
        com_tab.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                com_tabMouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                com_tabMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                com_tabMouseExited(evt);
            }
            public void mousePressed(java.awt.event.MouseEvent evt) {
                com_tabMousePressed(evt);
            }
        });
        com_tab.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jPanel21.setBackground(new java.awt.Color(228, 57, 39));
        jPanel21.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel56.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel56.setForeground(new java.awt.Color(255, 255, 255));
        jLabel56.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel56.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Home_1.png"))); // NOI18N
        jPanel21.add(jLabel56, new org.netbeans.lib.awtextra.AbsoluteConstraints(40, 10, 50, 40));

        jLabel57.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel57.setForeground(new java.awt.Color(255, 255, 255));
        jLabel57.setText("HOME");
        jPanel21.add(jLabel57, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        com_tab.add(jPanel21, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 60, 300, 60));

        jPanel22.setBackground(new java.awt.Color(228, 57, 39));
        jPanel22.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel58.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel58.setForeground(new java.awt.Color(255, 255, 255));
        jLabel58.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel58.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Home_1.png"))); // NOI18N
        jPanel22.add(jLabel58, new org.netbeans.lib.awtextra.AbsoluteConstraints(40, 10, 50, 40));

        jLabel59.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel59.setForeground(new java.awt.Color(255, 255, 255));
        jLabel59.setText("HOME");
        jPanel22.add(jLabel59, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        jPanel23.setBackground(new java.awt.Color(228, 57, 39));
        jPanel23.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel60.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel60.setForeground(new java.awt.Color(255, 255, 255));
        jLabel60.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel60.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Home_1.png"))); // NOI18N
        jPanel23.add(jLabel60, new org.netbeans.lib.awtextra.AbsoluteConstraints(40, 10, 50, 40));

        jLabel61.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel61.setForeground(new java.awt.Color(255, 255, 255));
        jLabel61.setText("HOME");
        jPanel23.add(jLabel61, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        jPanel22.add(jPanel23, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 60, 300, 60));

        com_tab.add(jPanel22, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 60, 300, 60));

        jLabel62.setFont(new java.awt.Font("Segoe UI", 0, 16)); // NOI18N
        jLabel62.setForeground(new java.awt.Color(228, 57, 39));
        jLabel62.setText("COMPUTER");
        jLabel62.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jLabel62MouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                jLabel62MouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                jLabel62MouseExited(evt);
            }
        });
        com_tab.add(jLabel62, new org.netbeans.lib.awtextra.AbsoluteConstraints(186, 5, -1, -1));

        jLabel55.setFont(new java.awt.Font("Segoe UI", 0, 16)); // NOI18N
        jLabel55.setForeground(new java.awt.Color(228, 57, 39));
        jLabel55.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Workstation_1.png"))); // NOI18N
        jLabel55.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jLabel55MouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                jLabel55MouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                jLabel55MouseExited(evt);
            }
        });
        com_tab.add(jLabel55, new org.netbeans.lib.awtextra.AbsoluteConstraints(130, 0, 50, 30));

        jPanel1.add(com_tab, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 400, 310, 30));

        jPanel5.setBackground(new java.awt.Color(228, 57, 39));
        jPanel5.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jPanel5MouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                jPanel5MouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                jPanel5MouseExited(evt);
            }
        });
        jPanel5.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel6.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel6.setForeground(new java.awt.Color(255, 255, 255));
        jLabel6.setText("LOG OUT");
        jPanel5.add(jLabel6, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        jLabel3.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel3.setForeground(new java.awt.Color(255, 255, 255));
        jLabel3.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel3.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Logout_1.png"))); // NOI18N
        jPanel5.add(jLabel3, new org.netbeans.lib.awtextra.AbsoluteConstraints(40, 0, 60, 60));

        jPanel1.add(jPanel5, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 560, 299, 60));

        inv_search_tab.setBackground(new java.awt.Color(228, 57, 39));
        inv_search_tab.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                inv_search_tabMouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                inv_search_tabMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                inv_search_tabMouseExited(evt);
            }
            public void mousePressed(java.awt.event.MouseEvent evt) {
                inv_search_tabMousePressed(evt);
            }
        });
        inv_search_tab.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel66.setFont(new java.awt.Font("Segoe UI", 0, 13)); // NOI18N
        jLabel66.setForeground(new java.awt.Color(228, 57, 39));
        jLabel66.setText("SEARCH INVENTORY");
        jLabel66.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jLabel66MouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                jLabel66MouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                jLabel66MouseExited(evt);
            }
        });
        inv_search_tab.add(jLabel66, new org.netbeans.lib.awtextra.AbsoluteConstraints(186, 8, -1, -1));

        jPanel24.setBackground(new java.awt.Color(228, 57, 39));
        jPanel24.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel67.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel67.setForeground(new java.awt.Color(255, 255, 255));
        jLabel67.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel67.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Home_1.png"))); // NOI18N
        jPanel24.add(jLabel67, new org.netbeans.lib.awtextra.AbsoluteConstraints(40, 10, 50, 40));

        jLabel68.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel68.setForeground(new java.awt.Color(255, 255, 255));
        jLabel68.setText("HOME");
        jPanel24.add(jLabel68, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        inv_search_tab.add(jPanel24, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 60, 300, 60));

        jPanel25.setBackground(new java.awt.Color(228, 57, 39));
        jPanel25.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel69.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel69.setForeground(new java.awt.Color(255, 255, 255));
        jLabel69.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel69.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Home_1.png"))); // NOI18N
        jPanel25.add(jLabel69, new org.netbeans.lib.awtextra.AbsoluteConstraints(40, 10, 50, 40));

        jLabel70.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel70.setForeground(new java.awt.Color(255, 255, 255));
        jLabel70.setText("HOME");
        jPanel25.add(jLabel70, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        jPanel26.setBackground(new java.awt.Color(228, 57, 39));
        jPanel26.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel71.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel71.setForeground(new java.awt.Color(255, 255, 255));
        jLabel71.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel71.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Home_1.png"))); // NOI18N
        jPanel26.add(jLabel71, new org.netbeans.lib.awtextra.AbsoluteConstraints(40, 10, 50, 40));

        jLabel72.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel72.setForeground(new java.awt.Color(255, 255, 255));
        jLabel72.setText("HOME");
        jPanel26.add(jLabel72, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        jPanel25.add(jPanel26, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 60, 300, 60));

        inv_search_tab.add(jPanel25, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 60, 300, 60));

        jLabel73.setFont(new java.awt.Font("Segoe UI", 0, 16)); // NOI18N
        jLabel73.setForeground(new java.awt.Color(228, 57, 39));
        jLabel73.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Search_7.png"))); // NOI18N
        jLabel73.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jLabel73MouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                jLabel73MouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                jLabel73MouseExited(evt);
            }
        });
        inv_search_tab.add(jLabel73, new org.netbeans.lib.awtextra.AbsoluteConstraints(130, 0, 50, 30));

        jPanel1.add(inv_search_tab, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 480, 310, 30));

        viber_tab.setBackground(new java.awt.Color(228, 57, 39));
        viber_tab.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                viber_tabMouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                viber_tabMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                viber_tabMouseExited(evt);
            }
            public void mousePressed(java.awt.event.MouseEvent evt) {
                viber_tabMousePressed(evt);
            }
        });
        viber_tab.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel89.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel89.setForeground(new java.awt.Color(255, 255, 255));
        jLabel89.setText("VIBER ACCOUNT");
        jLabel89.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jLabel89MouseClicked(evt);
            }
        });
        viber_tab.add(jLabel89, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        jLabel90.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel90.setForeground(new java.awt.Color(255, 255, 255));
        jLabel90.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel90.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/viber-19547.png"))); // NOI18N
        jLabel90.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jLabel90MouseClicked(evt);
            }
        });
        viber_tab.add(jLabel90, new org.netbeans.lib.awtextra.AbsoluteConstraints(30, 0, 60, 60));

        jPanel1.add(viber_tab, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 200, 310, 60));

        email_tab.setBackground(new java.awt.Color(228, 57, 39));
        email_tab.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                email_tabMouseClicked(evt);
            }
            public void mouseEntered(java.awt.event.MouseEvent evt) {
                email_tabMouseEntered(evt);
            }
            public void mouseExited(java.awt.event.MouseEvent evt) {
                email_tabMouseExited(evt);
            }
            public void mousePressed(java.awt.event.MouseEvent evt) {
                email_tabMousePressed(evt);
            }
        });
        email_tab.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel96.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel96.setForeground(new java.awt.Color(255, 255, 255));
        jLabel96.setText("EMAIL LIST");
        jLabel96.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jLabel96MouseClicked(evt);
            }
        });
        email_tab.add(jLabel96, new org.netbeans.lib.awtextra.AbsoluteConstraints(120, 20, -1, -1));

        jLabel97.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel97.setForeground(new java.awt.Color(255, 255, 255));
        jLabel97.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel97.setIcon(new javax.swing.ImageIcon(getClass().getResource("/picture/Email.png"))); // NOI18N
        jLabel97.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                jLabel97MouseClicked(evt);
            }
        });
        email_tab.add(jLabel97, new org.netbeans.lib.awtextra.AbsoluteConstraints(30, 0, 60, 60));

        jPanel1.add(email_tab, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 270, 310, 60));

        jPanel14.setBackground(new java.awt.Color(255, 255, 255));
        jPanel14.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel28.setFont(new java.awt.Font("Segoe UI", 1, 50)); // NOI18N
        jLabel28.setForeground(new java.awt.Color(255, 255, 255));
        jLabel28.setText("HOLDINGS INC.");
        jPanel14.add(jLabel28, new org.netbeans.lib.awtextra.AbsoluteConstraints(310, 75, 410, 50));

        jLabel120.setFont(new java.awt.Font("Segoe UI", 1, 120)); // NOI18N
        jLabel120.setForeground(new java.awt.Color(255, 255, 255));
        jLabel120.setText("JDC");
        jPanel14.add(jLabel120, new org.netbeans.lib.awtextra.AbsoluteConstraints(70, 10, 230, 130));

        jLabel121.setFont(new java.awt.Font("Segoe UI", 1, 50)); // NOI18N
        jLabel121.setForeground(new java.awt.Color(255, 255, 255));
        jLabel121.setText("GROUP");
        jPanel14.add(jLabel121, new org.netbeans.lib.awtextra.AbsoluteConstraints(310, 35, 180, 50));

        jLabel20.setIcon(new javax.swing.ImageIcon(getClass().getResource("/Untitled design.png"))); // NOI18N
        jPanel14.add(jLabel20, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 0, 1410, 150));

        jTabbedPane1.setBackground(new java.awt.Color(255, 255, 255));
        jTabbedPane1.setBorder(javax.swing.BorderFactory.createCompoundBorder());
        jTabbedPane1.setForeground(new java.awt.Color(255, 255, 255));
        jTabbedPane1.setTabLayoutPolicy(javax.swing.JTabbedPane.SCROLL_TAB_LAYOUT);
        jTabbedPane1.setEnabled(false);
        jTabbedPane1.setOpaque(true);

        jPanel4.setBackground(new java.awt.Color(255, 255, 255));
        jPanel4.setLayout(new java.awt.CardLayout());

        jPanel2.setBackground(new java.awt.Color(255, 255, 255));
        jPanel2.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel18.setFont(new java.awt.Font("Segoe UI", 1, 48)); // NOI18N
        jLabel18.setForeground(new java.awt.Color(255, 255, 255));
        jLabel18.setText("Welcome to JDC official software.");
        jPanel2.add(jLabel18, new org.netbeans.lib.awtextra.AbsoluteConstraints(150, 50, 780, 66));

        jLabel23.setIcon(new javax.swing.ImageIcon(getClass().getResource("/Cuts-of-Beef-3x2-1-a557f31f8b13462185b4f2c17ab5b746.png"))); // NOI18N
        jPanel2.add(jLabel23, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 0, 1110, 180));

        jLabel24.setIcon(new javax.swing.ImageIcon(getClass().getResource("/Screenshot 2023-08-17 131118.png"))); // NOI18N
        jLabel24.setText("jLabel24");
        jPanel2.add(jLabel24, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 370, 1100, 160));

        jLabel25.setFont(new java.awt.Font("Segoe UI", 1, 36)); // NOI18N
        jLabel25.setForeground(new java.awt.Color(255, 255, 255));
        jLabel25.setText("Welcome to JDC official software.");
        jPanel2.add(jLabel25, new org.netbeans.lib.awtextra.AbsoluteConstraints(140, 40, 597, 66));

        jLabel26.setFont(new java.awt.Font("Segoe UI", 3, 36)); // NOI18N
        jLabel26.setForeground(new java.awt.Color(0, 0, 153));
        jLabel26.setText("Umbrella Operation");
        jPanel2.add(jLabel26, new org.netbeans.lib.awtextra.AbsoluteConstraints(410, 300, 340, 50));

        jLabel27.setFont(new java.awt.Font("Segoe UI", 3, 36)); // NOI18N
        jLabel27.setForeground(new java.awt.Color(255, 0, 0));
        jLabel27.setText("Guaranteed Fresh from Our Farm!");
        jPanel2.add(jLabel27, new org.netbeans.lib.awtextra.AbsoluteConstraints(290, 220, 610, 50));

        jPanel4.add(jPanel2, "card2");

        jTabbedPane1.addTab("HOME", jPanel4);

        jPanel10.setBackground(new java.awt.Color(255, 255, 255));
        jPanel10.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        search_table.setAutoCreateRowSorter(true);
        search_table.setFont(new java.awt.Font("Microsoft JhengHei", 0, 14)); // NOI18N
        search_table.setModel(new javax.swing.table.DefaultTableModel(
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
        search_table.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        search_table.setFocusable(false);
        search_table.setGridColor(new java.awt.Color(255, 255, 255));
        search_table.setRowHeight(25);
        search_table.setSelectionBackground(new java.awt.Color(228, 57, 39));
        search_table.setSelectionForeground(new java.awt.Color(255, 255, 255));
        search_table.setShowHorizontalLines(true);
        search_table.getTableHeader().setReorderingAllowed(false);
        search_table.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                search_tableMouseClicked(evt);
            }
        });
        jScrollPane1.setViewportView(search_table);

        jPanel10.add(jScrollPane1, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 250, 1090, 248));

        jLabel5.setText("License Type:");
        jPanel10.add(jLabel5, new org.netbeans.lib.awtextra.AbsoluteConstraints(380, 10, -1, 27));

        user_code.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        user_code.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                user_codeActionPerformed(evt);
            }
        });
        jPanel10.add(user_code, new org.netbeans.lib.awtextra.AbsoluteConstraints(83, 19, 276, 25));

        jLabel16.setText("Name:");
        jPanel10.add(jLabel16, new org.netbeans.lib.awtextra.AbsoluteConstraints(43, 52, -1, 30));

        user_name.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel10.add(user_name, new org.netbeans.lib.awtextra.AbsoluteConstraints(83, 53, 276, 28));

        user_department.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel10.add(user_department, new org.netbeans.lib.awtextra.AbsoluteConstraints(83, 88, 276, 25));

        jLabel17.setText("Department:");
        jPanel10.add(jLabel17, new org.netbeans.lib.awtextra.AbsoluteConstraints(11, 88, -1, 24));

        LType.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "TYPE1", "TYPE2", "TYPE3", "TYPE4", "TYPE5", "TYPE6", " " }));
        jPanel10.add(LType, new org.netbeans.lib.awtextra.AbsoluteConstraints(390, 40, 318, -1));

        SaveDB2.setBackground(new java.awt.Color(228, 57, 39));
        SaveDB2.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        SaveDB2.setForeground(new java.awt.Color(255, 255, 255));
        SaveDB2.setText("Add");
        SaveDB2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                SaveDB2ActionPerformed(evt);
            }
        });
        jPanel10.add(SaveDB2, new org.netbeans.lib.awtextra.AbsoluteConstraints(770, 10, 298, 46));

        user_remove.setBackground(new java.awt.Color(228, 57, 39));
        user_remove.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        user_remove.setForeground(new java.awt.Color(255, 255, 255));
        user_remove.setText("Remove");
        user_remove.setEnabled(false);
        user_remove.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                user_removeActionPerformed(evt);
            }
        });
        jPanel10.add(user_remove, new org.netbeans.lib.awtextra.AbsoluteConstraints(680, 547, 160, 30));

        jLabel19.setText("Activity:");
        jPanel10.add(jLabel19, new org.netbeans.lib.awtextra.AbsoluteConstraints(35, 119, -1, 24));

        user_activity.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel10.add(user_activity, new org.netbeans.lib.awtextra.AbsoluteConstraints(83, 119, 276, 25));

        jLabel21.setText("Database:");
        jPanel10.add(jLabel21, new org.netbeans.lib.awtextra.AbsoluteConstraints(23, 150, -1, 24));

        user_database.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel10.add(user_database, new org.netbeans.lib.awtextra.AbsoluteConstraints(82, 150, 276, 25));

        jLabel22.setText("Remarks:");
        jPanel10.add(jLabel22, new org.netbeans.lib.awtextra.AbsoluteConstraints(390, 80, -1, 30));

        user_remark.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel10.add(user_remark, new org.netbeans.lib.awtextra.AbsoluteConstraints(440, 80, 265, 28));

        l_print2.setBackground(new java.awt.Color(228, 57, 39));
        l_print2.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        l_print2.setForeground(new java.awt.Color(255, 255, 255));
        l_print2.setText("PRINT");
        l_print2.setMaximumSize(new java.awt.Dimension(175, 27));
        l_print2.setMinimumSize(new java.awt.Dimension(175, 27));
        l_print2.setPreferredSize(new java.awt.Dimension(175, 27));
        l_print2.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                l_print2ActionPerformed(evt);
            }
        });
        jPanel10.add(l_print2, new org.netbeans.lib.awtextra.AbsoluteConstraints(250, 547, 160, 30));

        user_edit.setBackground(new java.awt.Color(228, 57, 39));
        user_edit.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        user_edit.setForeground(new java.awt.Color(255, 255, 255));
        user_edit.setText("Edit");
        user_edit.setEnabled(false);
        user_edit.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                user_editActionPerformed(evt);
            }
        });
        jPanel10.add(user_edit, new org.netbeans.lib.awtextra.AbsoluteConstraints(770, 70, 298, 46));

        jPanel39.setBackground(new java.awt.Color(228, 57, 39));

        javax.swing.GroupLayout jPanel39Layout = new javax.swing.GroupLayout(jPanel39);
        jPanel39.setLayout(jPanel39Layout);
        jPanel39Layout.setHorizontalGroup(
            jPanel39Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 1110, Short.MAX_VALUE)
        );
        jPanel39Layout.setVerticalGroup(
            jPanel39Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 10, Short.MAX_VALUE)
        );

        jPanel10.add(jPanel39, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 190, 1110, 10));

        user_search.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel10.add(user_search, new org.netbeans.lib.awtextra.AbsoluteConstraints(540, 210, 340, 30));

        SaveDB21.setBackground(new java.awt.Color(228, 57, 39));
        SaveDB21.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        SaveDB21.setForeground(new java.awt.Color(255, 255, 255));
        SaveDB21.setText("Search");
        SaveDB21.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                SaveDB21ActionPerformed(evt);
            }
        });
        jPanel10.add(SaveDB21, new org.netbeans.lib.awtextra.AbsoluteConstraints(910, 210, 140, 30));

        jLabel95.setText("User Code:");
        jPanel10.add(jLabel95, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 18, -1, 27));

        jTabbedPane1.addTab("SEARCH", jPanel10);

        jPanel11.setBackground(new java.awt.Color(255, 255, 255));
        jPanel11.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        user_table.setAutoCreateRowSorter(true);
        user_table.setFont(new java.awt.Font("Microsoft JhengHei", 0, 14)); // NOI18N
        user_table.setModel(new javax.swing.table.DefaultTableModel(
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
        user_table.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        user_table.setFocusable(false);
        user_table.setGridColor(new java.awt.Color(255, 255, 255));
        user_table.setRowHeight(25);
        user_table.setSelectionBackground(new java.awt.Color(255, 0, 0));
        user_table.setSelectionForeground(new java.awt.Color(255, 255, 255));
        user_table.getTableHeader().setReorderingAllowed(false);
        jScrollPane2.setViewportView(user_table);

        jPanel11.add(jScrollPane2, new org.netbeans.lib.awtextra.AbsoluteConstraints(5, 190, 1099, 393));

        jButton1.setBackground(new java.awt.Color(228, 57, 39));
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
        jPanel11.add(jButton1, new org.netbeans.lib.awtextra.AbsoluteConstraints(680, 20, 197, 54));

        export.setBackground(new java.awt.Color(228, 57, 39));
        export.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        export.setForeground(new java.awt.Color(255, 255, 255));
        export.setText("EXPORT TO EXCEL FILE");
        export.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                exportActionPerformed(evt);
            }
        });
        jPanel11.add(export, new org.netbeans.lib.awtextra.AbsoluteConstraints(680, 80, 197, 52));

        SaveDB.setBackground(new java.awt.Color(228, 57, 39));
        SaveDB.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        SaveDB.setForeground(new java.awt.Color(255, 255, 255));
        SaveDB.setText("SAVE");
        SaveDB.setMaximumSize(new java.awt.Dimension(175, 27));
        SaveDB.setMinimumSize(new java.awt.Dimension(175, 27));
        SaveDB.setPreferredSize(new java.awt.Dimension(175, 27));
        SaveDB.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                SaveDBActionPerformed(evt);
            }
        });
        jPanel11.add(SaveDB, new org.netbeans.lib.awtextra.AbsoluteConstraints(890, 20, 197, 52));

        jLabel29.setIcon(new javax.swing.ImageIcon(getClass().getResource("/Screenshot 2023-08-17 131118.png"))); // NOI18N
        jLabel29.setText("jLabel24");
        jPanel11.add(jLabel29, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 20, 650, 180));

        SaveDB6.setBackground(new java.awt.Color(228, 57, 39));
        SaveDB6.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        SaveDB6.setForeground(new java.awt.Color(255, 255, 255));
        SaveDB6.setText("PRINT");
        SaveDB6.setMaximumSize(new java.awt.Dimension(175, 27));
        SaveDB6.setMinimumSize(new java.awt.Dimension(175, 27));
        SaveDB6.setPreferredSize(new java.awt.Dimension(175, 27));
        SaveDB6.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                SaveDB6ActionPerformed(evt);
            }
        });
        jPanel11.add(SaveDB6, new org.netbeans.lib.awtextra.AbsoluteConstraints(890, 80, 197, 52));

        jTabbedPane1.addTab("SAP USER", jPanel11);

        jTabbedPane2.setTabLayoutPolicy(javax.swing.JTabbedPane.SCROLL_TAB_LAYOUT);
        jTabbedPane2.setEnabled(false);

        COMPUTER.setBackground(new java.awt.Color(255, 255, 255));
        COMPUTER.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel64.setIcon(new javax.swing.ImageIcon(getClass().getResource("/Screenshot 2023-08-17 131118.png"))); // NOI18N
        jLabel64.setText("jLabel24");
        COMPUTER.add(jLabel64, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 20, 640, 150));

        c_import.setBackground(new java.awt.Color(228, 57, 39));
        c_import.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        c_import.setForeground(new java.awt.Color(255, 255, 255));
        c_import.setText("IMPORT EXCEL FILE");
        c_import.setMaximumSize(new java.awt.Dimension(84, 32));
        c_import.setMinimumSize(new java.awt.Dimension(84, 32));
        c_import.setPreferredSize(new java.awt.Dimension(84, 32));
        c_import.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                c_importActionPerformed(evt);
            }
        });
        COMPUTER.add(c_import, new org.netbeans.lib.awtextra.AbsoluteConstraints(680, 20, 197, 54));

        c_export.setBackground(new java.awt.Color(228, 57, 39));
        c_export.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        c_export.setForeground(new java.awt.Color(255, 255, 255));
        c_export.setText("EXPORT TO EXCEL FILE");
        c_export.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                c_exportActionPerformed(evt);
            }
        });
        COMPUTER.add(c_export, new org.netbeans.lib.awtextra.AbsoluteConstraints(680, 80, 197, 52));

        c_save.setBackground(new java.awt.Color(228, 57, 39));
        c_save.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        c_save.setForeground(new java.awt.Color(255, 255, 255));
        c_save.setText("SAVE");
        c_save.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                c_saveActionPerformed(evt);
            }
        });
        COMPUTER.add(c_save, new org.netbeans.lib.awtextra.AbsoluteConstraints(890, 20, 197, 52));

        jScrollPane3.setBackground(new java.awt.Color(228, 57, 39));
        jScrollPane3.setBorder(null);

        computer_table.setAutoCreateRowSorter(true);
        computer_table.setFont(new java.awt.Font("Microsoft JhengHei", 0, 14)); // NOI18N
        computer_table.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "NAME", "DEPARTMENT", "MONITOR BRAND", "MONITOR ASSET BRAND", "MOUSE AND KEYBOARD", "MOTHERBOARD BRAND/MODEL", "MOTHERBOARD SERIAL NO.", "POWERSUPPLY BRAND", "POWER SUPPLY SERIAL NO.", "HARD DRIVE BRAND", "HARD DRIVE SIZE", "HARD DRIVE SERIAL NO.", "MEMORY BRAND", "MEMORY SIZE", "MEMORY SERIAL NO.", "GRAPHIC CARDS", "SERIAL NUMBER", "PROCESSOR", "PROCESSOR SPECS", "OFFICE LICENSE ACTIVATED", "WINDOWS LICENSE ACTIVATED", "IP ADDRESS", "YOUTUBE BLOCKED", "FB BLOCKED", "USB ENABLE?", "DOMAIN", "WEBCAM", "HEADSET", "WARRANTY END DATE PROCESSOR"
            }
        ) {
            boolean[] canEdit = new boolean [] {
                false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false, false
            };

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        computer_table.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        computer_table.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        computer_table.setFocusable(false);
        computer_table.setGridColor(new java.awt.Color(255, 255, 255));
        computer_table.setRowHeight(25);
        computer_table.setSelectionBackground(new java.awt.Color(228, 57, 39));
        computer_table.setSelectionForeground(new java.awt.Color(255, 255, 255));
        computer_table.getTableHeader().setReorderingAllowed(false);
        jScrollPane3.setViewportView(computer_table);
        if (computer_table.getColumnModel().getColumnCount() > 0) {
            computer_table.getColumnModel().getColumn(0).setResizable(false);
            computer_table.getColumnModel().getColumn(1).setResizable(false);
            computer_table.getColumnModel().getColumn(2).setResizable(false);
            computer_table.getColumnModel().getColumn(3).setResizable(false);
            computer_table.getColumnModel().getColumn(4).setResizable(false);
            computer_table.getColumnModel().getColumn(5).setResizable(false);
            computer_table.getColumnModel().getColumn(6).setResizable(false);
            computer_table.getColumnModel().getColumn(7).setResizable(false);
            computer_table.getColumnModel().getColumn(8).setResizable(false);
            computer_table.getColumnModel().getColumn(9).setResizable(false);
            computer_table.getColumnModel().getColumn(10).setResizable(false);
            computer_table.getColumnModel().getColumn(11).setResizable(false);
            computer_table.getColumnModel().getColumn(12).setResizable(false);
            computer_table.getColumnModel().getColumn(13).setResizable(false);
            computer_table.getColumnModel().getColumn(14).setResizable(false);
            computer_table.getColumnModel().getColumn(15).setResizable(false);
            computer_table.getColumnModel().getColumn(16).setResizable(false);
            computer_table.getColumnModel().getColumn(17).setResizable(false);
            computer_table.getColumnModel().getColumn(18).setResizable(false);
            computer_table.getColumnModel().getColumn(19).setResizable(false);
            computer_table.getColumnModel().getColumn(20).setResizable(false);
            computer_table.getColumnModel().getColumn(21).setResizable(false);
            computer_table.getColumnModel().getColumn(22).setResizable(false);
            computer_table.getColumnModel().getColumn(23).setResizable(false);
            computer_table.getColumnModel().getColumn(24).setResizable(false);
            computer_table.getColumnModel().getColumn(25).setResizable(false);
            computer_table.getColumnModel().getColumn(26).setResizable(false);
            computer_table.getColumnModel().getColumn(27).setResizable(false);
            computer_table.getColumnModel().getColumn(28).setResizable(false);
        }

        COMPUTER.add(jScrollPane3, new org.netbeans.lib.awtextra.AbsoluteConstraints(8, 190, 1090, 360));

        c_print.setBackground(new java.awt.Color(228, 57, 39));
        c_print.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        c_print.setForeground(new java.awt.Color(255, 255, 255));
        c_print.setText("PRINT");
        c_print.setMaximumSize(new java.awt.Dimension(175, 27));
        c_print.setMinimumSize(new java.awt.Dimension(175, 27));
        c_print.setPreferredSize(new java.awt.Dimension(175, 27));
        c_print.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                c_printActionPerformed(evt);
            }
        });
        COMPUTER.add(c_print, new org.netbeans.lib.awtextra.AbsoluteConstraints(890, 80, 197, 52));

        jTabbedPane2.addTab("COMPUTER", COMPUTER);

        jPanel3.setBackground(new java.awt.Color(255, 255, 255));
        jPanel3.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jScrollPane5.setBackground(new java.awt.Color(255, 204, 204));
        jScrollPane5.setBorder(null);

        laptop_table.setAutoCreateRowSorter(true);
        laptop_table.setBorder(javax.swing.BorderFactory.createEmptyBorder(1, 1, 1, 1));
        laptop_table.setFont(new java.awt.Font("Microsoft JhengHei", 0, 14)); // NOI18N
        laptop_table.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "DEPARTMENT", "ASSET ID", "ASSET DESCRIPTION", "BRAND", "MODEL", "SERIAL NUMBER", "ACCOUNTABLE TO", "WARRANTY DATE", "CONDITION", "STATUS", "RECOMMENDATION"
            }
        ) {
            Class[] types = new Class [] {
                java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class, java.lang.String.class
            };
            boolean[] canEdit = new boolean [] {
                false, false, false, false, false, false, false, false, false, true, false
            };

            public Class getColumnClass(int columnIndex) {
                return types [columnIndex];
            }

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        laptop_table.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        laptop_table.setFocusable(false);
        laptop_table.setGridColor(new java.awt.Color(255, 255, 255));
        laptop_table.setRowHeight(25);
        laptop_table.setSelectionBackground(new java.awt.Color(228, 57, 39));
        laptop_table.setSelectionForeground(new java.awt.Color(255, 255, 255));
        laptop_table.setShowHorizontalLines(true);
        laptop_table.getTableHeader().setReorderingAllowed(false);
        jScrollPane5.setViewportView(laptop_table);

        jPanel3.add(jScrollPane5, new org.netbeans.lib.awtextra.AbsoluteConstraints(6, 179, 1090, 373));

        jLabel65.setIcon(new javax.swing.ImageIcon(getClass().getResource("/Screenshot 2023-08-17 131118.png"))); // NOI18N
        jLabel65.setText("jLabel24");
        jPanel3.add(jLabel65, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 20, 640, 150));

        l_import.setBackground(new java.awt.Color(228, 57, 39));
        l_import.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        l_import.setForeground(new java.awt.Color(255, 255, 255));
        l_import.setText("IMPORT EXCEL FILE");
        l_import.setMaximumSize(new java.awt.Dimension(84, 32));
        l_import.setMinimumSize(new java.awt.Dimension(84, 32));
        l_import.setPreferredSize(new java.awt.Dimension(84, 32));
        l_import.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                l_importActionPerformed(evt);
            }
        });
        jPanel3.add(l_import, new org.netbeans.lib.awtextra.AbsoluteConstraints(680, 20, 197, 54));

        l_export.setBackground(new java.awt.Color(228, 57, 39));
        l_export.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        l_export.setForeground(new java.awt.Color(255, 255, 255));
        l_export.setText("EXPORT TO EXCEL FILE");
        l_export.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                l_exportActionPerformed(evt);
            }
        });
        jPanel3.add(l_export, new org.netbeans.lib.awtextra.AbsoluteConstraints(680, 80, 197, 52));

        l_save.setBackground(new java.awt.Color(228, 57, 39));
        l_save.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        l_save.setForeground(new java.awt.Color(255, 255, 255));
        l_save.setText("SAVE");
        l_save.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                l_saveActionPerformed(evt);
            }
        });
        jPanel3.add(l_save, new org.netbeans.lib.awtextra.AbsoluteConstraints(890, 20, 197, 52));

        l_print.setBackground(new java.awt.Color(228, 57, 39));
        l_print.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        l_print.setForeground(new java.awt.Color(255, 255, 255));
        l_print.setText("PRINT");
        l_print.setMaximumSize(new java.awt.Dimension(175, 27));
        l_print.setMinimumSize(new java.awt.Dimension(175, 27));
        l_print.setPreferredSize(new java.awt.Dimension(175, 27));
        l_print.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                l_printActionPerformed(evt);
            }
        });
        jPanel3.add(l_print, new org.netbeans.lib.awtextra.AbsoluteConstraints(890, 80, 197, 52));

        jTabbedPane2.addTab("LAPTOP", jPanel3);

        jPanel6.setBackground(new java.awt.Color(255, 255, 255));
        jPanel6.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        Assets.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Computer", "Laptop" }));
        Assets.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                AssetsActionPerformed(evt);
            }
        });
        jPanel6.add(Assets, new org.netbeans.lib.awtextra.AbsoluteConstraints(720, 30, 170, 30));

        SaveDB10.setBackground(new java.awt.Color(228, 57, 39));
        SaveDB10.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        SaveDB10.setForeground(new java.awt.Color(255, 255, 255));
        SaveDB10.setText("Search");
        SaveDB10.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                SaveDB10ActionPerformed(evt);
            }
        });
        jPanel6.add(SaveDB10, new org.netbeans.lib.awtextra.AbsoluteConstraints(900, 30, 140, 30));

        search.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        search.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel6.add(search, new org.netbeans.lib.awtextra.AbsoluteConstraints(360, 30, 340, 30));

        jScrollPane6.setBackground(new java.awt.Color(255, 204, 204));
        jScrollPane6.setBorder(null);

        inv_search_table.setAutoCreateRowSorter(true);
        inv_search_table.setBorder(javax.swing.BorderFactory.createEmptyBorder(1, 1, 1, 1));
        inv_search_table.setFont(new java.awt.Font("Microsoft JhengHei", 0, 14)); // NOI18N
        inv_search_table.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {

            }
        ));
        inv_search_table.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        inv_search_table.setFocusable(false);
        inv_search_table.setGridColor(new java.awt.Color(255, 255, 255));
        inv_search_table.setRowHeight(25);
        inv_search_table.setSelectionBackground(new java.awt.Color(228, 57, 39));
        inv_search_table.setSelectionForeground(new java.awt.Color(255, 255, 255));
        inv_search_table.setShowHorizontalLines(true);
        inv_search_table.getTableHeader().setReorderingAllowed(false);
        jScrollPane6.setViewportView(inv_search_table);

        jPanel6.add(jScrollPane6, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 110, 1090, 373));

        SaveDB9.setBackground(new java.awt.Color(228, 57, 39));
        SaveDB9.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        SaveDB9.setForeground(new java.awt.Color(255, 255, 255));
        SaveDB9.setText("Add");
        SaveDB9.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                SaveDB9ActionPerformed(evt);
            }
        });
        jPanel6.add(SaveDB9, new org.netbeans.lib.awtextra.AbsoluteConstraints(90, 500, 160, -1));

        SaveDB11.setBackground(new java.awt.Color(228, 57, 39));
        SaveDB11.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        SaveDB11.setForeground(new java.awt.Color(255, 255, 255));
        SaveDB11.setText("Remove");
        SaveDB11.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                SaveDB11ActionPerformed(evt);
            }
        });
        jPanel6.add(SaveDB11, new org.netbeans.lib.awtextra.AbsoluteConstraints(900, 500, 160, -1));

        l_print1.setBackground(new java.awt.Color(228, 57, 39));
        l_print1.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        l_print1.setForeground(new java.awt.Color(255, 255, 255));
        l_print1.setText("PRINT");
        l_print1.setMaximumSize(new java.awt.Dimension(175, 27));
        l_print1.setMinimumSize(new java.awt.Dimension(175, 27));
        l_print1.setPreferredSize(new java.awt.Dimension(175, 27));
        l_print1.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                l_print1ActionPerformed(evt);
            }
        });
        jPanel6.add(l_print1, new org.netbeans.lib.awtextra.AbsoluteConstraints(500, 500, 160, 30));

        jTabbedPane2.addTab("SEARCH INVENTORY", jPanel6);

        jPanel30.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jPanel31.setBackground(new java.awt.Color(255, 255, 255));
        jPanel31.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jPanel32.setBackground(new java.awt.Color(228, 57, 39));
        jPanel32.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());
        jPanel31.add(jPanel32, new org.netbeans.lib.awtextra.AbsoluteConstraints(660, 520, 340, 5));

        jPanel33.setBackground(new java.awt.Color(228, 57, 39));
        jPanel33.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());
        jPanel31.add(jPanel33, new org.netbeans.lib.awtextra.AbsoluteConstraints(150, 570, 500, 5));

        jPanel34.setBackground(new java.awt.Color(228, 57, 39));
        jPanel34.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());
        jPanel31.add(jPanel34, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 20, 5, 250));

        c_n.setToolTipText("");
        c_n.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel31.add(c_n, new org.netbeans.lib.awtextra.AbsoluteConstraints(200, 30, 276, 20));

        c_dept.setToolTipText("");
        c_dept.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel31.add(c_dept, new org.netbeans.lib.awtextra.AbsoluteConstraints(200, 60, 276, 20));

        c_mak.setToolTipText("");
        c_mak.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel31.add(c_mak, new org.netbeans.lib.awtextra.AbsoluteConstraints(200, 150, 276, 20));

        c_mn.setToolTipText("");
        c_mn.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel31.add(c_mn, new org.netbeans.lib.awtextra.AbsoluteConstraints(200, 90, 276, 20));

        c_mab.setToolTipText("");
        c_mab.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel31.add(c_mab, new org.netbeans.lib.awtextra.AbsoluteConstraints(200, 120, 276, 20));

        c_hds.setToolTipText("");
        c_hds.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel31.add(c_hds, new org.netbeans.lib.awtextra.AbsoluteConstraints(200, 330, 276, 20));

        c_pssn.setToolTipText("");
        c_pssn.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel31.add(c_pssn, new org.netbeans.lib.awtextra.AbsoluteConstraints(200, 270, 276, 20));

        c_psb.setToolTipText("");
        c_psb.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel31.add(c_psb, new org.netbeans.lib.awtextra.AbsoluteConstraints(200, 240, 276, 20));

        c_mbm.setToolTipText("");
        c_mbm.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel31.add(c_mbm, new org.netbeans.lib.awtextra.AbsoluteConstraints(200, 180, 276, 20));

        c_hdb.setToolTipText("");
        c_hdb.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel31.add(c_hdb, new org.netbeans.lib.awtextra.AbsoluteConstraints(200, 300, 276, 20));

        SaveDB13.setBackground(new java.awt.Color(228, 57, 39));
        SaveDB13.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        SaveDB13.setForeground(new java.awt.Color(255, 255, 255));
        SaveDB13.setText("Add");
        SaveDB13.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                SaveDB13ActionPerformed(evt);
            }
        });
        jPanel31.add(SaveDB13, new org.netbeans.lib.awtextra.AbsoluteConstraints(470, 510, 160, -1));

        jLabel78.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel78.setText("Department:");
        jPanel31.add(jLabel78, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 60, -1, -1));

        jLabel79.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel79.setText("Name:");
        jPanel31.add(jLabel79, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 30, -1, -1));

        jLabel80.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel80.setText("Monitor Asset Brand:");
        jPanel31.add(jLabel80, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 120, -1, -1));

        jLabel81.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel81.setText("Monitor Brand:");
        jPanel31.add(jLabel81, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 90, -1, -1));

        jLabel82.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel82.setText("Motherboard Brand/Model:");
        jPanel31.add(jLabel82, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 180, -1, -1));

        jLabel83.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel83.setText("Mouse and Keyboard:");
        jPanel31.add(jLabel83, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 150, -1, -1));

        jLabel84.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel84.setText("Hard Drive Size:");
        jPanel31.add(jLabel84, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 330, -1, -1));

        jLabel85.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel85.setText("Hard Drive Brand:");
        jPanel31.add(jLabel85, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 300, -1, -1));

        jLabel86.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel86.setText("Power Supply Serial No.:");
        jPanel31.add(jLabel86, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 270, -1, -1));

        jLabel87.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel87.setText("Power Supply Brand:");
        jPanel31.add(jLabel87, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 240, -1, -1));

        jLabel88.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel88.setText("Motherboard Serial No.:");
        jPanel31.add(jLabel88, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 210, -1, -1));

        c_ms.setToolTipText("");
        c_ms.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel31.add(c_ms, new org.netbeans.lib.awtextra.AbsoluteConstraints(200, 420, 276, 20));

        jLabel100.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel100.setText("Memory Size:");
        jPanel31.add(jLabel100, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 420, -1, -1));

        jLabel101.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel101.setText("Memory Brand:");
        jPanel31.add(jLabel101, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 390, -1, -1));

        c_mb.setToolTipText("");
        c_mb.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel31.add(c_mb, new org.netbeans.lib.awtextra.AbsoluteConstraints(200, 390, 276, 20));

        c_hdsn.setToolTipText("");
        c_hdsn.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel31.add(c_hdsn, new org.netbeans.lib.awtextra.AbsoluteConstraints(200, 360, 276, 20));

        jLabel102.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel102.setText("Hard Drive Serial No.:");
        jPanel31.add(jLabel102, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 360, -1, -1));

        jLabel103.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel103.setText("Memory Serial No.:");
        jPanel31.add(jLabel103, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 450, -1, -1));

        c_mmrysn.setToolTipText("");
        c_mmrysn.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel31.add(c_mmrysn, new org.netbeans.lib.awtextra.AbsoluteConstraints(200, 450, 276, 20));

        jLabel104.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel104.setText("Graphic Cards:");
        jPanel31.add(jLabel104, new org.netbeans.lib.awtextra.AbsoluteConstraints(550, 30, -1, -1));

        c_gc.setToolTipText("");
        c_gc.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel31.add(c_gc, new org.netbeans.lib.awtextra.AbsoluteConstraints(750, 30, 276, 20));

        jLabel105.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel105.setText("Serial Number:");
        jPanel31.add(jLabel105, new org.netbeans.lib.awtextra.AbsoluteConstraints(550, 60, -1, -1));

        c_sn.setToolTipText("");
        c_sn.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel31.add(c_sn, new org.netbeans.lib.awtextra.AbsoluteConstraints(750, 60, 276, 20));

        jLabel106.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel106.setText("Processor:");
        jPanel31.add(jLabel106, new org.netbeans.lib.awtextra.AbsoluteConstraints(550, 90, -1, -1));

        c_p.setToolTipText("");
        c_p.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel31.add(c_p, new org.netbeans.lib.awtextra.AbsoluteConstraints(750, 90, 276, 20));

        jLabel107.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel107.setText("Processor Specs:");
        jPanel31.add(jLabel107, new org.netbeans.lib.awtextra.AbsoluteConstraints(550, 120, -1, -1));

        c_ps.setToolTipText("");
        c_ps.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel31.add(c_ps, new org.netbeans.lib.awtextra.AbsoluteConstraints(750, 120, 276, 20));

        jLabel108.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel108.setText("Office License Activated:");
        jPanel31.add(jLabel108, new org.netbeans.lib.awtextra.AbsoluteConstraints(550, 150, -1, -1));

        jLabel109.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel109.setText("Windows License Activated:");
        jPanel31.add(jLabel109, new org.netbeans.lib.awtextra.AbsoluteConstraints(550, 180, -1, -1));

        jLabel110.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel110.setText("IP Address:");
        jPanel31.add(jLabel110, new org.netbeans.lib.awtextra.AbsoluteConstraints(550, 210, -1, -1));

        c_ip.setToolTipText("");
        c_ip.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel31.add(c_ip, new org.netbeans.lib.awtextra.AbsoluteConstraints(750, 210, 276, 20));

        jLabel111.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel111.setText("Youtube Blocked:");
        jPanel31.add(jLabel111, new org.netbeans.lib.awtextra.AbsoluteConstraints(550, 240, -1, -1));

        jLabel112.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel112.setText("FB Blocked:");
        jPanel31.add(jLabel112, new org.netbeans.lib.awtextra.AbsoluteConstraints(550, 270, -1, -1));

        jLabel113.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel113.setText("USB Enabled:");
        jPanel31.add(jLabel113, new org.netbeans.lib.awtextra.AbsoluteConstraints(550, 300, -1, -1));

        jLabel114.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel114.setText("Domain:");
        jPanel31.add(jLabel114, new org.netbeans.lib.awtextra.AbsoluteConstraints(550, 330, -1, -1));

        jLabel115.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel115.setText("Webcam:");
        jPanel31.add(jLabel115, new org.netbeans.lib.awtextra.AbsoluteConstraints(550, 360, -1, -1));

        jLabel116.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel116.setText("Headset:");
        jPanel31.add(jLabel116, new org.netbeans.lib.awtextra.AbsoluteConstraints(550, 390, -1, -1));

        jLabel117.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel117.setText("Warranty End Date Processor:");
        jPanel31.add(jLabel117, new org.netbeans.lib.awtextra.AbsoluteConstraints(550, 420, -1, -1));

        c_msn.setToolTipText("");
        c_msn.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel31.add(c_msn, new org.netbeans.lib.awtextra.AbsoluteConstraints(200, 210, 276, 20));

        jPanel35.setBackground(new java.awt.Color(228, 57, 39));
        jPanel35.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());
        jPanel31.add(jPanel35, new org.netbeans.lib.awtextra.AbsoluteConstraints(1080, 50, 5, 250));

        jPanel36.setBackground(new java.awt.Color(228, 57, 39));
        jPanel36.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());
        jPanel31.add(jPanel36, new org.netbeans.lib.awtextra.AbsoluteConstraints(100, 520, 340, 5));

        c_wla.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { " ", "Yes", "No" }));
        c_wla.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                c_wlaActionPerformed(evt);
            }
        });
        jPanel31.add(c_wla, new org.netbeans.lib.awtextra.AbsoluteConstraints(750, 180, 276, 20));

        c_oal.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { " ", "Yes", "No" }));
        c_oal.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                c_oalActionPerformed(evt);
            }
        });
        jPanel31.add(c_oal, new org.netbeans.lib.awtextra.AbsoluteConstraints(750, 150, 276, 20));

        c_yt.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { " ", "Yes", "No" }));
        c_yt.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                c_ytActionPerformed(evt);
            }
        });
        jPanel31.add(c_yt, new org.netbeans.lib.awtextra.AbsoluteConstraints(750, 240, 276, 20));

        c_fb.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { " ", "Yes", "No" }));
        c_fb.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                c_fbActionPerformed(evt);
            }
        });
        jPanel31.add(c_fb, new org.netbeans.lib.awtextra.AbsoluteConstraints(750, 270, 276, 20));

        c_usb.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { " ", "Yes", "No" }));
        c_usb.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                c_usbActionPerformed(evt);
            }
        });
        jPanel31.add(c_usb, new org.netbeans.lib.awtextra.AbsoluteConstraints(750, 300, 276, 20));

        c_d.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { " ", "Yes", "No" }));
        c_d.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                c_dActionPerformed(evt);
            }
        });
        jPanel31.add(c_d, new org.netbeans.lib.awtextra.AbsoluteConstraints(750, 330, 276, 20));

        c_h.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { " ", "Yes", "No" }));
        c_h.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                c_hActionPerformed(evt);
            }
        });
        jPanel31.add(c_h, new org.netbeans.lib.awtextra.AbsoluteConstraints(750, 390, 276, 20));

        c_w.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { " ", "Yes", "No" }));
        c_w.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                c_wActionPerformed(evt);
            }
        });
        jPanel31.add(c_w, new org.netbeans.lib.awtextra.AbsoluteConstraints(750, 360, 276, 20));
        jPanel31.add(c_wedp, new org.netbeans.lib.awtextra.AbsoluteConstraints(750, 420, 280, -1));

        jPanel30.add(jPanel31, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 0, 1100, 560));

        jTabbedPane2.addTab("", jPanel30);

        jPanel7.setBackground(new java.awt.Color(255, 255, 255));

        jPanel8.setBackground(new java.awt.Color(255, 255, 255));
        jPanel8.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jPanel27.setBackground(new java.awt.Color(228, 57, 39));
        jPanel27.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());
        jPanel8.add(jPanel27, new org.netbeans.lib.awtextra.AbsoluteConstraints(640, 120, 5, 250));

        jPanel28.setBackground(new java.awt.Color(228, 57, 39));
        jPanel28.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());
        jPanel8.add(jPanel28, new org.netbeans.lib.awtextra.AbsoluteConstraints(150, 570, 500, 5));

        jPanel29.setBackground(new java.awt.Color(228, 57, 39));
        jPanel29.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());
        jPanel8.add(jPanel29, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 20, 5, 250));

        dept1.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        dept1.setToolTipText("");
        dept1.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel8.add(dept1, new org.netbeans.lib.awtextra.AbsoluteConstraints(290, 70, 276, 28));

        id.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        id.setToolTipText("");
        id.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel8.add(id, new org.netbeans.lib.awtextra.AbsoluteConstraints(290, 110, 276, 28));

        mdl.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        mdl.setToolTipText("");
        mdl.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel8.add(mdl, new org.netbeans.lib.awtextra.AbsoluteConstraints(290, 230, 276, 28));

        desc.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        desc.setToolTipText("");
        desc.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel8.add(desc, new org.netbeans.lib.awtextra.AbsoluteConstraints(290, 150, 276, 28));

        brnd.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        brnd.setToolTipText("");
        brnd.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel8.add(brnd, new org.netbeans.lib.awtextra.AbsoluteConstraints(290, 190, 276, 28));

        reco.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        reco.setToolTipText("");
        reco.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel8.add(reco, new org.netbeans.lib.awtextra.AbsoluteConstraints(290, 470, 276, 28));

        condi.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        condi.setToolTipText("");
        condi.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel8.add(condi, new org.netbeans.lib.awtextra.AbsoluteConstraints(290, 390, 276, 28));

        acct.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        acct.setToolTipText("");
        acct.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel8.add(acct, new org.netbeans.lib.awtextra.AbsoluteConstraints(290, 310, 276, 28));

        srl.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        srl.setToolTipText("");
        srl.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel8.add(srl, new org.netbeans.lib.awtextra.AbsoluteConstraints(290, 270, 276, 28));

        status.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        status.setToolTipText("");
        status.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        jPanel8.add(status, new org.netbeans.lib.awtextra.AbsoluteConstraints(290, 430, 276, 28));

        SaveDB12.setBackground(new java.awt.Color(228, 57, 39));
        SaveDB12.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        SaveDB12.setForeground(new java.awt.Color(255, 255, 255));
        SaveDB12.setText("Add");
        SaveDB12.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                SaveDB12ActionPerformed(evt);
            }
        });
        jPanel8.add(SaveDB12, new org.netbeans.lib.awtextra.AbsoluteConstraints(250, 520, 160, -1));

        jLabel7.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel7.setText("Asset ID:");
        jPanel8.add(jLabel7, new org.netbeans.lib.awtextra.AbsoluteConstraints(130, 110, -1, -1));

        jLabel8.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel8.setText("Department:");
        jPanel8.add(jLabel8, new org.netbeans.lib.awtextra.AbsoluteConstraints(130, 70, -1, -1));

        jLabel9.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel9.setText("Brand:");
        jPanel8.add(jLabel9, new org.netbeans.lib.awtextra.AbsoluteConstraints(130, 190, -1, -1));

        jLabel10.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel10.setText("Asset Description:");
        jPanel8.add(jLabel10, new org.netbeans.lib.awtextra.AbsoluteConstraints(130, 150, -1, -1));

        jLabel11.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel11.setText("Serial Number:");
        jPanel8.add(jLabel11, new org.netbeans.lib.awtextra.AbsoluteConstraints(130, 270, -1, -1));

        jLabel12.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel12.setText("Model:");
        jPanel8.add(jLabel12, new org.netbeans.lib.awtextra.AbsoluteConstraints(130, 230, -1, -1));

        jLabel13.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel13.setText("Recommendation:");
        jPanel8.add(jLabel13, new org.netbeans.lib.awtextra.AbsoluteConstraints(130, 470, -1, -1));

        jLabel74.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel74.setText("Status:");
        jPanel8.add(jLabel74, new org.netbeans.lib.awtextra.AbsoluteConstraints(130, 430, -1, -1));

        jLabel75.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel75.setText("Conditions:");
        jPanel8.add(jLabel75, new org.netbeans.lib.awtextra.AbsoluteConstraints(130, 390, -1, -1));

        jLabel76.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel76.setText("Warranty Date:");
        jPanel8.add(jLabel76, new org.netbeans.lib.awtextra.AbsoluteConstraints(130, 350, -1, -1));

        jLabel77.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        jLabel77.setText("Accountable to:");
        jPanel8.add(jLabel77, new org.netbeans.lib.awtextra.AbsoluteConstraints(130, 310, -1, -1));
        jPanel8.add(date, new org.netbeans.lib.awtextra.AbsoluteConstraints(290, 350, 280, -1));

        javax.swing.GroupLayout jPanel7Layout = new javax.swing.GroupLayout(jPanel7);
        jPanel7.setLayout(jPanel7Layout);
        jPanel7Layout.setHorizontalGroup(
            jPanel7Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 1102, Short.MAX_VALUE)
            .addGroup(jPanel7Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addGroup(jPanel7Layout.createSequentialGroup()
                    .addGap(0, 0, Short.MAX_VALUE)
                    .addComponent(jPanel8, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGap(0, 0, Short.MAX_VALUE)))
        );
        jPanel7Layout.setVerticalGroup(
            jPanel7Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 590, Short.MAX_VALUE)
            .addGroup(jPanel7Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                .addGroup(jPanel7Layout.createSequentialGroup()
                    .addGap(0, 0, Short.MAX_VALUE)
                    .addComponent(jPanel8, javax.swing.GroupLayout.PREFERRED_SIZE, 590, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGap(0, 0, Short.MAX_VALUE)))
        );

        jTabbedPane2.addTab("", jPanel7);

        jTabbedPane1.addTab("INVENTORY", jTabbedPane2);

        jPanel37.setBackground(new java.awt.Color(255, 255, 255));
        jPanel37.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel91.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel91.setText("Device Type:");
        jPanel37.add(jLabel91, new org.netbeans.lib.awtextra.AbsoluteConstraints(500, 30, -1, -1));

        client_name.setToolTipText("");
        client_name.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        client_name.addKeyListener(new java.awt.event.KeyAdapter() {
            public void keyReleased(java.awt.event.KeyEvent evt) {
                client_nameKeyReleased(evt);
            }
            public void keyTyped(java.awt.event.KeyEvent evt) {
                client_nameKeyTyped(evt);
            }
        });
        jPanel37.add(client_name, new org.netbeans.lib.awtextra.AbsoluteConstraints(200, 30, 276, 20));

        jLabel92.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel92.setText("Mobile Number:");
        jPanel37.add(jLabel92, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 60, -1, -1));

        mobile_number.setToolTipText("");
        mobile_number.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        mobile_number.addKeyListener(new java.awt.event.KeyAdapter() {
            public void keyReleased(java.awt.event.KeyEvent evt) {
                mobile_numberKeyReleased(evt);
            }
            public void keyTyped(java.awt.event.KeyEvent evt) {
                mobile_numberKeyTyped(evt);
            }
        });
        jPanel37.add(mobile_number, new org.netbeans.lib.awtextra.AbsoluteConstraints(200, 60, 276, 20));

        jLabel93.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel93.setText("Department:");
        jPanel37.add(jLabel93, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 90, -1, -1));

        department.setToolTipText("");
        department.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        department.addKeyListener(new java.awt.event.KeyAdapter() {
            public void keyReleased(java.awt.event.KeyEvent evt) {
                departmentKeyReleased(evt);
            }
            public void keyTyped(java.awt.event.KeyEvent evt) {
                departmentKeyTyped(evt);
            }
        });
        jPanel37.add(department, new org.netbeans.lib.awtextra.AbsoluteConstraints(200, 90, 276, 20));

        SaveDB14.setBackground(new java.awt.Color(228, 57, 39));
        SaveDB14.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        SaveDB14.setForeground(new java.awt.Color(255, 255, 255));
        SaveDB14.setText("Print");
        SaveDB14.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                SaveDB14ActionPerformed(evt);
            }
        });
        jPanel37.add(SaveDB14, new org.netbeans.lib.awtextra.AbsoluteConstraints(910, 125, 160, -1));

        Viber_add.setBackground(new java.awt.Color(228, 57, 39));
        Viber_add.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        Viber_add.setForeground(new java.awt.Color(255, 255, 255));
        Viber_add.setText("Add");
        Viber_add.setEnabled(false);
        Viber_add.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                Viber_addActionPerformed(evt);
            }
        });
        jPanel37.add(Viber_add, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 130, 160, -1));

        Export_Viber.setBackground(new java.awt.Color(228, 57, 39));
        Export_Viber.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        Export_Viber.setForeground(new java.awt.Color(255, 255, 255));
        Export_Viber.setText("Export");
        Export_Viber.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                Export_ViberActionPerformed(evt);
            }
        });
        jPanel37.add(Export_Viber, new org.netbeans.lib.awtextra.AbsoluteConstraints(380, 125, 160, -1));

        Delete.setBackground(new java.awt.Color(228, 57, 39));
        Delete.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        Delete.setForeground(new java.awt.Color(255, 255, 255));
        Delete.setText("Delete");
        Delete.setEnabled(false);
        Delete.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                DeleteActionPerformed(evt);
            }
        });
        jPanel37.add(Delete, new org.netbeans.lib.awtextra.AbsoluteConstraints(550, 125, 160, -1));

        Viber_edit.setBackground(new java.awt.Color(228, 57, 39));
        Viber_edit.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        Viber_edit.setForeground(new java.awt.Color(255, 255, 255));
        Viber_edit.setText("Edit");
        Viber_edit.setEnabled(false);
        Viber_edit.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                Viber_editActionPerformed(evt);
            }
        });
        jPanel37.add(Viber_edit, new org.netbeans.lib.awtextra.AbsoluteConstraints(730, 127, 160, 30));

        jScrollPane7.setBackground(new java.awt.Color(255, 204, 204));
        jScrollPane7.setBorder(null);

        viber_table.setAutoCreateRowSorter(true);
        viber_table.setBorder(javax.swing.BorderFactory.createEmptyBorder(1, 1, 1, 1));
        viber_table.setFont(new java.awt.Font("Microsoft JhengHei", 0, 14)); // NOI18N
        viber_table.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "Client Name:", "Mobile Number:", "Department:", "Device Type:"
            }
        ) {
            boolean[] canEdit = new boolean [] {
                false, false, false, false
            };

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        viber_table.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        viber_table.setFocusable(false);
        viber_table.setGridColor(new java.awt.Color(255, 255, 255));
        viber_table.setRowHeight(25);
        viber_table.setSelectionBackground(new java.awt.Color(228, 57, 39));
        viber_table.setSelectionForeground(new java.awt.Color(255, 255, 255));
        viber_table.setShowHorizontalLines(true);
        viber_table.getTableHeader().setReorderingAllowed(false);
        viber_table.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                viber_tableMouseClicked(evt);
            }
        });
        jScrollPane7.setViewportView(viber_table);
        if (viber_table.getColumnModel().getColumnCount() > 0) {
            viber_table.getColumnModel().getColumn(0).setResizable(false);
            viber_table.getColumnModel().getColumn(1).setResizable(false);
            viber_table.getColumnModel().getColumn(2).setResizable(false);
            viber_table.getColumnModel().getColumn(3).setResizable(false);
        }

        jPanel37.add(jScrollPane7, new org.netbeans.lib.awtextra.AbsoluteConstraints(5, 260, 1090, 280));

        Viber_Import.setBackground(new java.awt.Color(228, 57, 39));
        Viber_Import.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        Viber_Import.setForeground(new java.awt.Color(255, 255, 255));
        Viber_Import.setText("Import");
        Viber_Import.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                Viber_ImportActionPerformed(evt);
            }
        });
        jPanel37.add(Viber_Import, new org.netbeans.lib.awtextra.AbsoluteConstraints(200, 125, 160, -1));
        jPanel37.add(viber_search, new org.netbeans.lib.awtextra.AbsoluteConstraints(560, 210, 340, 30));

        SaveDB20.setBackground(new java.awt.Color(228, 57, 39));
        SaveDB20.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        SaveDB20.setForeground(new java.awt.Color(255, 255, 255));
        SaveDB20.setText("Search");
        SaveDB20.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                SaveDB20ActionPerformed(evt);
            }
        });
        jPanel37.add(SaveDB20, new org.netbeans.lib.awtextra.AbsoluteConstraints(930, 210, 140, 30));

        jPanel38.setBackground(new java.awt.Color(228, 57, 39));

        javax.swing.GroupLayout jPanel38Layout = new javax.swing.GroupLayout(jPanel38);
        jPanel38.setLayout(jPanel38Layout);
        jPanel38Layout.setHorizontalGroup(
            jPanel38Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 1105, Short.MAX_VALUE)
        );
        jPanel38Layout.setVerticalGroup(
            jPanel38Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 10, Short.MAX_VALUE)
        );

        jPanel37.add(jPanel38, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 170, 1105, 10));

        device_type.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        device_type.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Phone only", "Desktop only", "Phone and Desktop" }));
        device_type.addItemListener(new java.awt.event.ItemListener() {
            public void itemStateChanged(java.awt.event.ItemEvent evt) {
                device_typeItemStateChanged(evt);
            }
        });
        device_type.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseReleased(java.awt.event.MouseEvent evt) {
                device_typeMouseReleased(evt);
            }
        });
        device_type.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                device_typeActionPerformed(evt);
            }
        });
        jPanel37.add(device_type, new org.netbeans.lib.awtextra.AbsoluteConstraints(590, 28, 300, 25));

        jLabel94.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel94.setText("Client Name:");
        jPanel37.add(jLabel94, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 30, -1, -1));

        jTabbedPane1.addTab("VIBER", jPanel37);

        jPanel40.setBackground(new java.awt.Color(255, 255, 255));
        jPanel40.setLayout(new org.netbeans.lib.awtextra.AbsoluteLayout());

        jLabel98.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel98.setText("Name:");
        jPanel40.add(jLabel98, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 30, -1, -1));

        email_name.setToolTipText("");
        email_name.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        email_name.addKeyListener(new java.awt.event.KeyAdapter() {
            public void keyReleased(java.awt.event.KeyEvent evt) {
                email_nameKeyReleased(evt);
            }
            public void keyTyped(java.awt.event.KeyEvent evt) {
                email_nameKeyTyped(evt);
            }
        });
        jPanel40.add(email_name, new org.netbeans.lib.awtextra.AbsoluteConstraints(100, 30, 350, 20));

        jLabel99.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel99.setText("Position:");
        jPanel40.add(jLabel99, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 60, -1, -1));

        email_position.setToolTipText("");
        email_position.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        email_position.addKeyListener(new java.awt.event.KeyAdapter() {
            public void keyReleased(java.awt.event.KeyEvent evt) {
                email_positionKeyReleased(evt);
            }
            public void keyTyped(java.awt.event.KeyEvent evt) {
                email_positionKeyTyped(evt);
            }
        });
        jPanel40.add(email_position, new org.netbeans.lib.awtextra.AbsoluteConstraints(100, 60, 350, 20));

        jLabel118.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel118.setText("Email:");
        jPanel40.add(jLabel118, new org.netbeans.lib.awtextra.AbsoluteConstraints(20, 90, -1, -1));

        email_email.setToolTipText("");
        email_email.setBorder(javax.swing.BorderFactory.createLineBorder(new java.awt.Color(0, 0, 0)));
        email_email.addKeyListener(new java.awt.event.KeyAdapter() {
            public void keyReleased(java.awt.event.KeyEvent evt) {
                email_emailKeyReleased(evt);
            }
            public void keyTyped(java.awt.event.KeyEvent evt) {
                email_emailKeyTyped(evt);
            }
        });
        jPanel40.add(email_email, new org.netbeans.lib.awtextra.AbsoluteConstraints(100, 90, 350, 20));

        jLabel119.setFont(new java.awt.Font("Segoe UI", 0, 14)); // NOI18N
        jLabel119.setText("Department:");
        jPanel40.add(jLabel119, new org.netbeans.lib.awtextra.AbsoluteConstraints(500, 30, -1, -1));

        email_department.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "HR Department", "Purchasing Department", "Accounting Department", "Audit Department", "Admin Department", "IT Department" }));
        email_department.addItemListener(new java.awt.event.ItemListener() {
            public void itemStateChanged(java.awt.event.ItemEvent evt) {
                email_departmentItemStateChanged(evt);
            }
        });
        email_department.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseReleased(java.awt.event.MouseEvent evt) {
                email_departmentMouseReleased(evt);
            }
        });
        email_department.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                email_departmentActionPerformed(evt);
            }
        });
        jPanel40.add(email_department, new org.netbeans.lib.awtextra.AbsoluteConstraints(600, 30, 276, 20));

        email_add.setBackground(new java.awt.Color(228, 57, 39));
        email_add.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        email_add.setForeground(new java.awt.Color(255, 255, 255));
        email_add.setText("Add");
        email_add.setEnabled(false);
        email_add.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                email_addActionPerformed(evt);
            }
        });
        jPanel40.add(email_add, new org.netbeans.lib.awtextra.AbsoluteConstraints(10, 125, 160, -1));

        email_import.setBackground(new java.awt.Color(228, 57, 39));
        email_import.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        email_import.setForeground(new java.awt.Color(255, 255, 255));
        email_import.setText("Import");
        email_import.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                email_importActionPerformed(evt);
            }
        });
        jPanel40.add(email_import, new org.netbeans.lib.awtextra.AbsoluteConstraints(200, 125, 160, -1));

        email_export.setBackground(new java.awt.Color(228, 57, 39));
        email_export.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        email_export.setForeground(new java.awt.Color(255, 255, 255));
        email_export.setText("Export");
        email_export.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                email_exportActionPerformed(evt);
            }
        });
        jPanel40.add(email_export, new org.netbeans.lib.awtextra.AbsoluteConstraints(380, 125, 160, -1));

        email_delete.setBackground(new java.awt.Color(228, 57, 39));
        email_delete.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        email_delete.setForeground(new java.awt.Color(255, 255, 255));
        email_delete.setText("Delete");
        email_delete.setEnabled(false);
        email_delete.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                email_deleteActionPerformed(evt);
            }
        });
        jPanel40.add(email_delete, new org.netbeans.lib.awtextra.AbsoluteConstraints(550, 125, 160, -1));

        email_edit.setBackground(new java.awt.Color(228, 57, 39));
        email_edit.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        email_edit.setForeground(new java.awt.Color(255, 255, 255));
        email_edit.setText("Edit");
        email_edit.setEnabled(false);
        email_edit.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                email_editActionPerformed(evt);
            }
        });
        jPanel40.add(email_edit, new org.netbeans.lib.awtextra.AbsoluteConstraints(730, 127, 160, 30));

        email_print.setBackground(new java.awt.Color(228, 57, 39));
        email_print.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        email_print.setForeground(new java.awt.Color(255, 255, 255));
        email_print.setText("Print");
        email_print.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                email_printActionPerformed(evt);
            }
        });
        jPanel40.add(email_print, new org.netbeans.lib.awtextra.AbsoluteConstraints(910, 125, 160, -1));

        jPanel41.setBackground(new java.awt.Color(228, 57, 39));

        javax.swing.GroupLayout jPanel41Layout = new javax.swing.GroupLayout(jPanel41);
        jPanel41.setLayout(jPanel41Layout);
        jPanel41Layout.setHorizontalGroup(
            jPanel41Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 1105, Short.MAX_VALUE)
        );
        jPanel41Layout.setVerticalGroup(
            jPanel41Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGap(0, 10, Short.MAX_VALUE)
        );

        jPanel40.add(jPanel41, new org.netbeans.lib.awtextra.AbsoluteConstraints(0, 170, 1105, 10));
        jPanel40.add(email_search, new org.netbeans.lib.awtextra.AbsoluteConstraints(560, 210, 340, 30));

        email_searchbutton.setBackground(new java.awt.Color(228, 57, 39));
        email_searchbutton.setFont(new java.awt.Font("Segoe UI", 0, 18)); // NOI18N
        email_searchbutton.setForeground(new java.awt.Color(255, 255, 255));
        email_searchbutton.setText("Search");
        email_searchbutton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                email_searchbuttonActionPerformed(evt);
            }
        });
        jPanel40.add(email_searchbutton, new org.netbeans.lib.awtextra.AbsoluteConstraints(920, 210, 140, 30));

        jScrollPane8.setBackground(new java.awt.Color(255, 204, 204));
        jScrollPane8.setBorder(null);

        email_table.setAutoCreateRowSorter(true);
        email_table.setBorder(javax.swing.BorderFactory.createEmptyBorder(1, 1, 1, 1));
        email_table.setFont(new java.awt.Font("Microsoft JhengHei", 0, 14)); // NOI18N
        email_table.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "Name", "Position", "Department:", "Email"
            }
        ) {
            boolean[] canEdit = new boolean [] {
                false, false, false, false
            };

            public boolean isCellEditable(int rowIndex, int columnIndex) {
                return canEdit [columnIndex];
            }
        });
        email_table.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        email_table.setFocusable(false);
        email_table.setGridColor(new java.awt.Color(255, 255, 255));
        email_table.setRowHeight(25);
        email_table.setSelectionBackground(new java.awt.Color(228, 57, 39));
        email_table.setSelectionForeground(new java.awt.Color(255, 255, 255));
        email_table.setShowHorizontalLines(true);
        email_table.getTableHeader().setReorderingAllowed(false);
        email_table.addMouseListener(new java.awt.event.MouseAdapter() {
            public void mouseClicked(java.awt.event.MouseEvent evt) {
                email_tableMouseClicked(evt);
            }
        });
        jScrollPane8.setViewportView(email_table);
        if (email_table.getColumnModel().getColumnCount() > 0) {
            email_table.getColumnModel().getColumn(0).setResizable(false);
            email_table.getColumnModel().getColumn(1).setResizable(false);
            email_table.getColumnModel().getColumn(2).setResizable(false);
            email_table.getColumnModel().getColumn(3).setResizable(false);
        }

        jPanel40.add(jScrollPane8, new org.netbeans.lib.awtextra.AbsoluteConstraints(5, 260, 1090, 280));

        jTabbedPane1.addTab("EMAIL", jPanel40);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, 308, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(0, 0, 0)
                .addComponent(jTabbedPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 0, Short.MAX_VALUE))
            .addComponent(jPanel14, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
            .addGroup(layout.createSequentialGroup()
                .addGap(0, 0, 0)
                .addComponent(jPanel14, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(0, 0, 0)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jTabbedPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 0, Short.MAX_VALUE)
                    .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)))
        );

        pack();
        setLocationRelativeTo(null);
    }// </editor-fold>//GEN-END:initComponents
    public void setColor(JPanel p) {
        p.setBackground(new Color(195, 0, 0));
    }

    public void resetColor(JPanel p) {
        p.setBackground(new Color(228, 57, 39));
    }

    public void setTable(JTable t) {
        t.getTableHeader().setFont(new Font("Segoe UI", Font.BOLD, 14));
        t.getTableHeader().setBackground(new Color(228, 57, 39));
        t.getTableHeader().setForeground(new Color(0, 0, 0));
        t.getTableHeader().setOpaque(true);
        t.getColumnModel().getColumn(0).setPreferredWidth(200);
        t.getColumnModel().getColumn(1).setPreferredWidth(200);
        t.getColumnModel().getColumn(2).setPreferredWidth(200);
        t.getColumnModel().getColumn(3).setPreferredWidth(200);
        t.getColumnModel().getColumn(4).setPreferredWidth(200);
        t.getColumnModel().getColumn(5).setPreferredWidth(200);
        t.getColumnModel().getColumn(6).setPreferredWidth(200);
    }

    public void ViberTable(JTable t) {
        t.getTableHeader().setFont(new Font("Segoe UI", Font.BOLD, 14));
        t.getTableHeader().setBackground(new Color(228, 57, 39));
        t.getTableHeader().setForeground(new Color(0, 0, 0));
        t.getTableHeader().setOpaque(true);
        t.getColumnModel().getColumn(0).setPreferredWidth(273);
        t.getColumnModel().getColumn(1).setPreferredWidth(273);
        t.getColumnModel().getColumn(2).setPreferredWidth(272);
        t.getColumnModel().getColumn(3).setPreferredWidth(272);
    }

    public void setInventory(JTable t) {
        t.getTableHeader().setFont(new Font("Segoe UI", Font.BOLD, 14));
        t.getTableHeader().setBackground(new Color(228, 57, 39));
        t.getTableHeader().setForeground(new Color(0, 0, 0));
        t.getTableHeader().setOpaque(true);
        t.getColumnModel().getColumn(0).setPreferredWidth(200);
        t.getColumnModel().getColumn(1).setPreferredWidth(200);
        t.getColumnModel().getColumn(2).setPreferredWidth(200);
        t.getColumnModel().getColumn(3).setPreferredWidth(200);
        t.getColumnModel().getColumn(4).setPreferredWidth(200);
        t.getColumnModel().getColumn(5).setPreferredWidth(200);
        t.getColumnModel().getColumn(6).setPreferredWidth(200);
        t.getColumnModel().getColumn(7).setPreferredWidth(200);
        t.getColumnModel().getColumn(8).setPreferredWidth(200);
        t.getColumnModel().getColumn(9).setPreferredWidth(200);
        t.getColumnModel().getColumn(10).setPreferredWidth(200);
    }

    public void setCInventory(JTable t) {
        t.getTableHeader().setFont(new Font("Segoe UI", Font.BOLD, 14));
        t.getTableHeader().setBackground(new Color(228, 57, 39));
        t.getTableHeader().setForeground(new Color(0, 0, 0));
        t.getTableHeader().setOpaque(true);
        t.getColumnModel().getColumn(0).setPreferredWidth(200);
        t.getColumnModel().getColumn(1).setPreferredWidth(200);
        t.getColumnModel().getColumn(2).setPreferredWidth(200);
        t.getColumnModel().getColumn(3).setPreferredWidth(200);
        t.getColumnModel().getColumn(4).setPreferredWidth(200);
        t.getColumnModel().getColumn(5).setPreferredWidth(200);
        t.getColumnModel().getColumn(6).setPreferredWidth(200);
        t.getColumnModel().getColumn(7).setPreferredWidth(200);
        t.getColumnModel().getColumn(8).setPreferredWidth(200);
        t.getColumnModel().getColumn(9).setPreferredWidth(200);
        t.getColumnModel().getColumn(10).setPreferredWidth(200);
        t.getColumnModel().getColumn(11).setPreferredWidth(200);
        t.getColumnModel().getColumn(12).setPreferredWidth(200);
        t.getColumnModel().getColumn(13).setPreferredWidth(200);
        t.getColumnModel().getColumn(14).setPreferredWidth(200);
        t.getColumnModel().getColumn(15).setPreferredWidth(200);
        t.getColumnModel().getColumn(16).setPreferredWidth(200);
        t.getColumnModel().getColumn(17).setPreferredWidth(200);
        t.getColumnModel().getColumn(18).setPreferredWidth(200);
        t.getColumnModel().getColumn(19).setPreferredWidth(200);
        t.getColumnModel().getColumn(20).setPreferredWidth(200);
        t.getColumnModel().getColumn(21).setPreferredWidth(200);
        t.getColumnModel().getColumn(22).setPreferredWidth(200);
        t.getColumnModel().getColumn(23).setPreferredWidth(200);
        t.getColumnModel().getColumn(24).setPreferredWidth(200);
        t.getColumnModel().getColumn(25).setPreferredWidth(200);
        t.getColumnModel().getColumn(26).setPreferredWidth(200);
        t.getColumnModel().getColumn(27).setPreferredWidth(200);
        t.getColumnModel().getColumn(28).setPreferredWidth(200);

    }

    public void settingtable() {
        String s = Assets.getSelectedItem().toString();
        DefaultTableModel df = (DefaultTableModel) inv_search_table.getModel();

        if (s == ("Computer")) {
            df.addColumn("NAME");
            df.addColumn("DEPARTMENT");
            df.addColumn("MONITOR BRAND");
            df.addColumn("MONITOR ASSET BRAND");
            df.addColumn("MOUSE AND KEYBOARD");
            df.addColumn("MOTHERBOARD BRAND/MODEL");
            df.addColumn("MOTHERBOARD SERIAL NO.");
            df.addColumn("POWERSUPPLY BRAND");
            df.addColumn("POWERSUPPLY SERIAL NO.");
            df.addColumn("HARD DRIVE BRAND");
            df.addColumn("HARD DRIVE SIZE");
            df.addColumn("HARD DRIVE SERIAL NO.");
            df.addColumn("MEMORY BRAND");
            df.addColumn("MEMORY SIZE");
            df.addColumn("MEMORY SERIAL NO.");
            df.addColumn("GRAPHIC CARDS");
            df.addColumn("SERIAL NUMBER");
            df.addColumn("PROCESSOR");
            df.addColumn("PROCESSOR SPECS");
            df.addColumn("OFFICE LICENSE ACTIVATED");
            df.addColumn("WINDOWS LICENSE ACTIVATED");
            df.addColumn("IP ADDRESS");
            df.addColumn("YOUTUBE BLOCKED");
            df.addColumn("FB BLOCKED");
            df.addColumn("USB ENABLED");
            df.addColumn("DOMAIN");
            df.addColumn("WEBCAM");
            df.addColumn("HEADSET");
            df.addColumn("WARRANTY END DATE PROCESSOR");
            setCInventory(inv_search_table);
        } else {
            df.addColumn("DEPARTMENT");
            df.addColumn("ASSET ID");
            df.addColumn("ASSET DESCRIPTION");
            df.addColumn("BRAND");
            df.addColumn("MODEL");
            df.addColumn("SERIAL NUMBER");
            df.addColumn("ACCOUNTABLE TO");
            df.addColumn("WARRANTY DATE");
            df.addColumn("CONDITION");
            df.addColumn("STATUS");
            df.addColumn("RECOMMENDATION");
            setInventory(inv_search_table);
        }

    }


    private void user_tabMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_user_tabMouseEntered

    }//GEN-LAST:event_user_tabMouseEntered

    private void user_tabMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_user_tabMouseExited

    }//GEN-LAST:event_user_tabMouseExited

    private void jPanel5MouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jPanel5MouseEntered

    }//GEN-LAST:event_jPanel5MouseEntered

    private void jPanel5MouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jPanel5MouseExited

    }//GEN-LAST:event_jPanel5MouseExited

    private void user_tabMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_user_tabMouseClicked
        jTabbedPane1.setSelectedIndex(2);
        resetColor(email_tab);
        setColor(user_tab);
        resetColor(search_tab);
        resetColor(home_tab);
        resetColor(inventory1);
        resetColor(com_tab);
        resetColor(lap_tab);
        resetColor(inv_search_tab);
        resetColor(viber_tab);
        setTable(user_table);
        jLabel62.setForeground(new Color(228, 57, 39));
        jLabel48.setForeground(new Color(228, 57, 39));
        jLabel66.setForeground(new Color(228, 57, 39));
        jLabel55.setVisible(false);
        jLabel63.setVisible(false);
        jLabel73.setVisible(false);
        inv = "false";

    }//GEN-LAST:event_user_tabMouseClicked

    private void home_tabMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_home_tabMouseClicked

        jTabbedPane1.setSelectedIndex(0);
        resetColor(email_tab);
        setColor(home_tab);
        resetColor(search_tab);
        resetColor(user_tab);
        resetColor(inventory1);
        resetColor(com_tab);
        resetColor(lap_tab);
        resetColor(viber_tab);
        jLabel62.setForeground(new Color(228, 57, 39));
        jLabel48.setForeground(new Color(228, 57, 39));
        jLabel66.setForeground(new Color(228, 57, 39));
        jLabel55.setVisible(false);
        jLabel63.setVisible(false);
        jLabel73.setVisible(false);
        inv = "false";


    }//GEN-LAST:event_home_tabMouseClicked

    private void home_tabMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_home_tabMouseEntered


    }//GEN-LAST:event_home_tabMouseEntered

    private void home_tabMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_home_tabMouseExited


    }//GEN-LAST:event_home_tabMouseExited

    private void jPanel5MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jPanel5MouseClicked
        int result = JOptionPane.showConfirmDialog(null, "Sure? You want to log out?", "Swing Tester",
                JOptionPane.YES_NO_OPTION, JOptionPane.INFORMATION_MESSAGE);
        if (result == JOptionPane.YES_OPTION) {
            Login log = new Login();
            log.setVisible(true);
            this.dispose();
        } else if (result == JOptionPane.NO_OPTION) {

        }

    }//GEN-LAST:event_jPanel5MouseClicked

    private void user_tabMousePressed(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_user_tabMousePressed

    }//GEN-LAST:event_user_tabMousePressed

    private void home_p1MouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_home_p1MouseExited
        // TODO add your handling code here:
    }//GEN-LAST:event_home_p1MouseExited

    private void home_p1MouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_home_p1MouseEntered
        // TODO add your handling code here:
    }//GEN-LAST:event_home_p1MouseEntered

    private void home_p1MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_home_p1MouseClicked
        // TODO add your handling code here:
    }//GEN-LAST:event_home_p1MouseClicked

    private void jLabel30MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel30MouseClicked

        jTabbedPane1.setSelectedIndex(1);
        resetColor(email_tab);
        setTable(search_table);
        setColor(search_tab);
        resetColor(home_tab);
        resetColor(user_tab);
        resetColor(inventory1);
        resetColor(com_tab);
        resetColor(lap_tab);
        resetColor(inv_search_tab);
        resetColor(viber_tab);
        jLabel62.setForeground(new Color(228, 57, 39));
        jLabel48.setForeground(new Color(228, 57, 39));
        jLabel66.setForeground(new Color(228, 57, 39));
        jLabel55.setVisible(false);
        jLabel63.setVisible(false);
        jLabel73.setVisible(false);
        user_code.requestFocus();
        com_tab.setEnabled(false);
        inv = "false";
        isEmpty();

    }//GEN-LAST:event_jLabel30MouseClicked

    private void jLabel30MouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel30MouseEntered
        // TODO add your handling code here:
    }//GEN-LAST:event_jLabel30MouseEntered

    private void jLabel30MouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel30MouseExited
        // TODO add your handling code here:
    }//GEN-LAST:event_jLabel30MouseExited

    private void search_tabMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_search_tabMouseClicked

        jTabbedPane1.setSelectedIndex(1);
        resetColor(email_tab);
        setTable(search_table);
        setColor(search_tab);
        resetColor(home_tab);
        resetColor(user_tab);
        resetColor(inventory1);
        resetColor(com_tab);
        resetColor(lap_tab);
        resetColor(inv_search_tab);
        resetColor(viber_tab);
        jLabel62.setForeground(new Color(228, 57, 39));
        jLabel48.setForeground(new Color(228, 57, 39));
        jLabel66.setForeground(new Color(228, 57, 39));
        jLabel55.setVisible(false);
        jLabel63.setVisible(false);
        jLabel73.setVisible(false);
        user_code.requestFocus();
        user_search.requestFocus();
        com_tab.setEnabled(false);
        inv = "false";
        isEmpty();
    }//GEN-LAST:event_search_tabMouseClicked

    private void search_tabMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_search_tabMouseEntered
        // TODO add your handling code here:
    }//GEN-LAST:event_search_tabMouseEntered

    private void search_tabMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_search_tabMouseExited
        // TODO add your handling code here:
    }//GEN-LAST:event_search_tabMouseExited

    private void search_tabMousePressed(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_search_tabMousePressed
        // TODO add your handling code here:
    }//GEN-LAST:event_search_tabMousePressed

    private void jLabel40MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel40MouseClicked
        jTabbedPane1.setSelectedIndex(3);
        jTabbedPane2.setSelectedIndex(0);
        setColor(inventory1);
        setColor(com_tab);
        resetColor(home_tab);
        resetColor(email_tab);
        resetColor(user_tab);
        resetColor(search_tab);
        resetColor(viber_tab);
        setTable(computer_table);
        setCInventory(computer_table);

        jLabel62.setForeground(Color.white);
        jLabel48.setForeground(Color.white);
        jLabel66.setForeground(Color.white);
        jLabel55.setVisible(true);
        jLabel63.setVisible(true);
        jLabel73.setVisible(true);
        Icon computer = jLabel55.getIcon();
        ImageIcon iconc = (ImageIcon) computer;
        Image imagec = iconc.getImage().getScaledInstance(jLabel55.getWidth(), jLabel55.getHeight(), Image.SCALE_SMOOTH);
        jLabel55.setIcon(new ImageIcon(imagec));

        Icon laptop = jLabel63.getIcon();
        ImageIcon iconl = (ImageIcon) laptop;
        Image imagel = iconl.getImage().getScaledInstance(jLabel63.getWidth(), jLabel63.getHeight(), Image.SCALE_SMOOTH);
        jLabel63.setIcon(new ImageIcon(imagel));
        inv = "true";

        Icon s = jLabel73.getIcon();
        ImageIcon icons = (ImageIcon) s;
        Image images = icons.getImage().getScaledInstance(jLabel73.getWidth(), jLabel73.getHeight(), Image.SCALE_SMOOTH);
        jLabel73.setIcon(new ImageIcon(images));
        inv = "true";
    }//GEN-LAST:event_jLabel40MouseClicked

    private void jLabel40MouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel40MouseEntered
        // TODO add your handling code here:
    }//GEN-LAST:event_jLabel40MouseEntered

    private void jLabel40MouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel40MouseExited
        // TODO add your handling code here:
    }//GEN-LAST:event_jLabel40MouseExited

    private void inventory1MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_inventory1MouseClicked
        jTabbedPane1.setSelectedIndex(3);
        jTabbedPane2.setSelectedIndex(0);
        setColor(inventory1);
        setColor(com_tab);
        resetColor(email_tab);
        resetColor(home_tab);
        resetColor(user_tab);
        resetColor(search_tab);
        resetColor(viber_tab);
        setTable(computer_table);
        setCInventory(computer_table);

        jLabel62.setForeground(Color.white);
        jLabel48.setForeground(Color.white);
        jLabel66.setForeground(Color.white);
        jLabel55.setVisible(true);
        jLabel63.setVisible(true);
        jLabel73.setVisible(true);
        Icon computer = jLabel55.getIcon();
        ImageIcon iconc = (ImageIcon) computer;
        Image imagec = iconc.getImage().getScaledInstance(jLabel55.getWidth(), jLabel55.getHeight(), Image.SCALE_SMOOTH);
        jLabel55.setIcon(new ImageIcon(imagec));

        Icon laptop = jLabel63.getIcon();
        ImageIcon iconl = (ImageIcon) laptop;
        Image imagel = iconl.getImage().getScaledInstance(jLabel63.getWidth(), jLabel63.getHeight(), Image.SCALE_SMOOTH);
        jLabel63.setIcon(new ImageIcon(imagel));
        inv = "true";

        Icon s = jLabel73.getIcon();
        ImageIcon icons = (ImageIcon) s;
        Image images = icons.getImage().getScaledInstance(jLabel73.getWidth(), jLabel73.getHeight(), Image.SCALE_SMOOTH);
        jLabel73.setIcon(new ImageIcon(images));
        inv = "true";
    }//GEN-LAST:event_inventory1MouseClicked

    private void inventory1MouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_inventory1MouseEntered
        // TODO add your handling code here:
    }//GEN-LAST:event_inventory1MouseEntered

    private void inventory1MouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_inventory1MouseExited
        // TODO add your handling code here:
    }//GEN-LAST:event_inventory1MouseExited

    private void inventory1MousePressed(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_inventory1MousePressed
        // TODO add your handling code here:
    }//GEN-LAST:event_inventory1MousePressed

    private void jLabel48MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel48MouseClicked
        if (inv == "true") {
            setColor(lap_tab);
            resetColor(com_tab);
            resetColor(inv_search_tab);
            resetColor(home_tab);
            resetColor(user_tab);
            resetColor(search_tab);
            resetColor(viber_tab);
            jTabbedPane2.setSelectedIndex(1);
            setTable(laptop_table);
            setInventory(laptop_table);
        } else {
            resetColor(lap_tab);
        }
    }//GEN-LAST:event_jLabel48MouseClicked

    private void jLabel48MouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel48MouseEntered
        // TODO add your handling code here:
    }//GEN-LAST:event_jLabel48MouseEntered

    private void jLabel48MouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel48MouseExited
        // TODO add your handling code here:
    }//GEN-LAST:event_jLabel48MouseExited

    private void lap_tabMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_lap_tabMouseClicked
        if (inv == "true") {
            setColor(lap_tab);
            resetColor(com_tab);
            resetColor(inv_search_tab);
            resetColor(home_tab);
            resetColor(user_tab);
            resetColor(search_tab);
            resetColor(viber_tab);
            jTabbedPane2.setSelectedIndex(1);
            setTable(laptop_table);
            setInventory(laptop_table);
        } else {
            resetColor(lap_tab);
        }
    }//GEN-LAST:event_lap_tabMouseClicked

    private void lap_tabMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_lap_tabMouseEntered
        // TODO add your handling code here:
    }//GEN-LAST:event_lap_tabMouseEntered

    private void lap_tabMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_lap_tabMouseExited
        // TODO add your handling code here:
    }//GEN-LAST:event_lap_tabMouseExited

    private void lap_tabMousePressed(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_lap_tabMousePressed
        // TODO add your handling code here:
    }//GEN-LAST:event_lap_tabMousePressed

    private void jLabel55MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel55MouseClicked
        if (inv == "true") {
            setColor(com_tab);
            resetColor(lap_tab);
            resetColor(inv_search_tab);
            resetColor(home_tab);
            resetColor(user_tab);
            resetColor(search_tab);
            resetColor(viber_tab);
            jTabbedPane2.setSelectedIndex(0);
            setTable(computer_table);
            setCInventory(computer_table);
        } else {
            resetColor(com_tab);
        }


    }//GEN-LAST:event_jLabel55MouseClicked

    private void jLabel55MouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel55MouseEntered
        // TODO add your handling code here:
    }//GEN-LAST:event_jLabel55MouseEntered

    private void jLabel55MouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel55MouseExited
        // TODO add your handling code here:
    }//GEN-LAST:event_jLabel55MouseExited

    private void com_tabMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_com_tabMouseClicked
        if (inv == "true") {
            setColor(com_tab);
            resetColor(lap_tab);
            resetColor(inv_search_tab);
            resetColor(home_tab);
            resetColor(user_tab);
            resetColor(search_tab);
            resetColor(viber_tab);
            jTabbedPane2.setSelectedIndex(0);
            setTable(computer_table);
            setCInventory(computer_table);
        } else {
            resetColor(com_tab);
        }


    }//GEN-LAST:event_com_tabMouseClicked

    private void com_tabMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_com_tabMouseEntered
        // TODO add your handling code here:
    }//GEN-LAST:event_com_tabMouseEntered

    private void com_tabMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_com_tabMouseExited
        // TODO add your handling code here:
    }//GEN-LAST:event_com_tabMouseExited

    private void com_tabMousePressed(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_com_tabMousePressed
        // TODO add your handling code here:
    }//GEN-LAST:event_com_tabMousePressed

    private void jLabel62MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel62MouseClicked
        if (inv == "true") {
            setColor(com_tab);
            resetColor(lap_tab);
            resetColor(inv_search_tab);
            resetColor(home_tab);
            resetColor(user_tab);
            resetColor(search_tab);
            resetColor(viber_tab);
            jTabbedPane2.setSelectedIndex(0);
            setTable(computer_table);
            setCInventory(computer_table);
        } else {
            resetColor(com_tab);
        }


    }//GEN-LAST:event_jLabel62MouseClicked

    private void jLabel62MouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel62MouseEntered
        // TODO add your handling code here:
    }//GEN-LAST:event_jLabel62MouseEntered

    private void jLabel62MouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel62MouseExited
        // TODO add your handling code here:
    }//GEN-LAST:event_jLabel62MouseExited

    private void jLabel63MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel63MouseClicked
        if (inv == "true") {
            setColor(lap_tab);
            resetColor(com_tab);
            resetColor(inv_search_tab);
            resetColor(home_tab);
            resetColor(user_tab);
            resetColor(search_tab);
            resetColor(viber_tab);
            jTabbedPane2.setSelectedIndex(1);
            setTable(laptop_table);
            setInventory(laptop_table);
        } else {
            resetColor(lap_tab);
        }
    }//GEN-LAST:event_jLabel63MouseClicked

    private void jLabel63MouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel63MouseEntered
        // TODO add your handling code here:
    }//GEN-LAST:event_jLabel63MouseEntered

    private void jLabel63MouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel63MouseExited
        // TODO add your handling code here:
    }//GEN-LAST:event_jLabel63MouseExited

    private void SaveDBActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_SaveDBActionPerformed
        int columnIndex = 3;
        DefaultTableModel model = (DefaultTableModel) user_table.getModel();
        int rowCount = model.getRowCount();
        int saved = 0;
        int dup = 1;
        if (rowCount != 0) {
            for (int row = 0; row < rowCount; row++) {
                Object value = model.getValueAt(row, columnIndex);
                try {

                    String duplicate = "SELECT * FROM user WHERE user_code = '" + value + "'";
                    pst = con.prepareStatement(duplicate);
                    rs = pst.executeQuery(duplicate);

                    if (rs.next()) {
                        if (dup == 1) {
                            dup = 0;
                            JOptionPane.showMessageDialog(this, "Some User Code already used ");
                        }

                    } else if (!rs.next()) {

                        int rowCount1 = user_table.getRowCount();
                        int columnCount = user_table.getColumnCount();
                        for (int row1 = 0; row1 < rowCount1; row1++) {

                            Object val = user_table.getValueAt(row1, 3);
                            Object dept = user_table.getValueAt(row1, 0);
                            Object act = user_table.getValueAt(row1, 1);
                            Object db = user_table.getValueAt(row1, 2);
                            Object user_code = user_table.getValueAt(row1, 3);
                            Object name = user_table.getValueAt(row1, 4);
                            Object license = user_table.getValueAt(row1, 5);
                            Object remarks = user_table.getValueAt(row1, 6);

                            if (dept == null) {
                                dept = " ";
                            }
                            if (act == null) {
                                act = " ";
                            }
                            if (db == null) {
                                db = " ";
                            }
                            if (user_code == null) {
                                user_code = " ";
                            }
                            if (name == null) {
                                name = " ";
                            }
                            if (license == null) {
                                license = " ";
                            }
                            if (remarks == null) {
                                remarks = " ";
                            }

                            if (val == value) {

                                pst = con.prepareStatement("Insert into user(department,activity,db,user_code,name,license,remarks) values (?,?,?,?,?,?,?)");
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
                    Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);
                }

            }
            if (saved == 1) {
                JOptionPane.showMessageDialog(null, "SAVED!");
            }
        } else {
            JOptionPane.showMessageDialog(null, "You need to import files");
        }
    }//GEN-LAST:event_SaveDBActionPerformed

    private void exportActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_exportActionPerformed
       
          DefaultTableModel model = (DefaultTableModel) user_table.getModel();
        int rowCount = model.getRowCount();
        if (rowCount == 0) {
            JOptionPane.showMessageDialog(null, "You need to import files");
        } else {
            try {
                String defaultCurrentDirectoryPath = "C:\\Users\\Arnie D\\Downloads\\Arnie\\OJT\\IT DEPT";
                JFileChooser jFileChooser = new JFileChooser(defaultCurrentDirectoryPath);
                jFileChooser.showSaveDialog(this);
                File saveFile = jFileChooser.getSelectedFile();
                if (saveFile != null) {
                    saveFile = new File(saveFile.toString() + ".xlsx");
                    if (saveFile.exists()) {
                        int response = JOptionPane.showConfirmDialog(null, //
                                "Do you want to replace the existing file?", //
                                "Confirm", JOptionPane.YES_NO_OPTION, //
                                JOptionPane.QUESTION_MESSAGE);
                        if (response != JOptionPane.YES_OPTION) {

                            jFileChooser.showSaveDialog(this);
                             saveFile = jFileChooser.getSelectedFile();
                            saveFile = new File(saveFile.toString()+ ".xlsx");
                            Workbook wb = new XSSFWorkbook();
                            Sheet sheet = wb.createSheet("Computer Inventory");
                            Row rowCol = sheet.createRow(0);
                            for (int c = 0; c < user_table.getColumnCount(); c++) {
                                Cell cell = rowCol.createCell(c);
                                cell.setCellValue(user_table.getColumnName(c));
                            }
                            for (int r = 0; r < user_table.getRowCount(); r++) {
                                Row row = sheet.createRow(r);
                                for (int rc = 0; rc < user_table.getColumnCount(); rc++) {
                                    Cell cell = row.createCell(rc);
                                    if (user_table.getValueAt(r, rc) != null) {
                                        cell.setCellValue(user_table.getValueAt(r, rc).toString());
                                    }
                                }
                            }
                            FileOutputStream out = new FileOutputStream(new File(saveFile.toString()));
                            wb.write(out);
                            wb.close();
                            out.close();
                            openFile(saveFile.toString());

                        }
                    } else if(!saveFile.exists()) {
                        Workbook wb = new XSSFWorkbook();
                        Sheet sheet = wb.createSheet("EMAIL LIST ACCOUNTS");
                        Row rowCol = sheet.createRow(0);
                        for (int c = 0; c < user_table.getColumnCount(); c++) {
                            Cell cell = rowCol.createCell(c);
                            cell.setCellValue(user_table.getColumnName(c));
                        }
                        for (int r = 0; r < user_table.getRowCount(); r++) {
                            Row row = sheet.createRow(r);
                            for (int rc = 0; rc < user_table.getColumnCount(); rc++) {
                                Cell cell = row.createCell(rc);
                                if (user_table.getValueAt(r, rc) != null) {
                                    cell.setCellValue(user_table.getValueAt(r, rc).toString());
                                }
                            }
                        }
                        FileOutputStream out = new FileOutputStream(new File(saveFile.toString()));
                        wb.write(out);
                        wb.close();
                        out.close();
                        openFile(saveFile.toString());
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Error!");
                }

            } catch (FileNotFoundException e) {
                System.out.println(e);
            } catch (IOException io) {
                System.out.println(io);
            }  
        }
    }//GEN-LAST:event_exportActionPerformed

    private void jButton1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton1ActionPerformed
        int RowCount = user_table.getRowCount();
        if (RowCount != 0) {
            DefaultTableModel df = (DefaultTableModel) user_table.getModel();
            df.setRowCount(0);
            RowCount = 0;
            if (RowCount == 0) {
                Import();
            }
        } else {
            Import();
        }
    }//GEN-LAST:event_jButton1ActionPerformed

    private void user_removeActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_user_removeActionPerformed
        try {
            DefaultTableModel model = (DefaultTableModel) search_table.getModel();
            if (search_table.getSelectedRowCount() == 0) {
                JOptionPane.showMessageDialog(this, "User not selected");
            } else {
                int row = search_table.getSelectedRow();
                String selected = (String) search_table.getModel().getValueAt(row, 3);
                pst = con.prepareStatement("DELETE FROM user WHERE user_code=?");
                pst.setString(1, selected);
                int k = pst.executeUpdate();
                if (k == 1) {
                    JOptionPane.showMessageDialog(this, "User deleted");
                    Fetch();
                    user_remove.setEnabled(false);
                    remove();
                } else {
                    JOptionPane.showMessageDialog(this, "User not deleted");
                }
            }
        } catch (SQLException ex) {
            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_user_removeActionPerformed

    private void SaveDB2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_SaveDB2ActionPerformed
        int saved = 0;
        try {
            String code = user_code.getText();
            String nm = user_name.getText();
            String dp = user_department.getText();
            String act = user_activity.getText();
            String db1 = user_database.getText();
            String lt = LType.getSelectedItem().toString();
            String r = user_remark.getText();
            if (code.equals("") || (nm.equals("") || (dp.equals("")
                    || (lt.equals("License Type") || (act.equals("") || (db1.equals("")) || (r.equals(""))))))) {
                JOptionPane.showMessageDialog(null, "Please input all the necessary information");
                user_code.requestFocus();
            } else {
                String dup = "Select * from user where user_code='" + code + "'and department='" + dp + "' and name='" + nm + "'";
                pst = con.prepareStatement(dup);
                rs = pst.executeQuery(dup);

                if (rs.next()) {
                    JOptionPane.showMessageDialog(null, "DUPLICATE");
                } else {
                    pst = con.prepareStatement("Insert into user(department,activity,db,user_code,name,license,remarks) values (?,?,?,?,?,?,?)");
                    pst.setString(1, dp);
                    pst.setString(2, act);
                    pst.setString(3, db1);
                    pst.setString(4, code);
                    pst.setString(5, nm);
                    pst.setString(6, lt);
                    pst.setString(7, r);
                    saved = pst.executeUpdate();

                }
                if (saved == 1) {
                    JOptionPane.showMessageDialog(null, "SAVED!");
                    user_code.setText("");
                    user_name.setText("");
                    user_department.setText("");
                    user_activity.setText("");
                    user_database.setText("");
                    LType.setSelectedIndex(0);
                    user_remark.setText("");

                    Fetch();

                }
            }

        } catch (SQLException ex) {
            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);

        }

    }//GEN-LAST:event_SaveDB2ActionPerformed

    private void user_codeActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_user_codeActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_user_codeActionPerformed

    private void jLabel66MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel66MouseClicked
        if (inv == "true") {
            setColor(inv_search_tab);
            resetColor(com_tab);
            resetColor(lap_tab);
            resetColor(home_tab);
            resetColor(user_tab);
            resetColor(search_tab);
            resetColor(viber_tab);
            jTabbedPane2.setSelectedIndex(2);
            Assets.setSelectedIndex(0);
            Fetch1();
            search.requestFocus();
        } else {
            resetColor(inv_search_tab);
        }
    }//GEN-LAST:event_jLabel66MouseClicked

    private void jLabel66MouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel66MouseEntered
        // TODO add your handling code here:
    }//GEN-LAST:event_jLabel66MouseEntered

    private void jLabel66MouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel66MouseExited
        // TODO add your handling code here:
    }//GEN-LAST:event_jLabel66MouseExited

    private void jLabel73MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel73MouseClicked
        if (inv == "true") {
            setColor(inv_search_tab);
            resetColor(com_tab);
            resetColor(lap_tab);
            resetColor(home_tab);
            resetColor(user_tab);
            resetColor(search_tab);
            resetColor(viber_tab);
            jTabbedPane2.setSelectedIndex(2);
            Assets.setSelectedIndex(0);
            Fetch1();
            search.requestFocus();
        } else {
            resetColor(inv_search_tab);
        }
    }//GEN-LAST:event_jLabel73MouseClicked

    private void jLabel73MouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel73MouseEntered
        // TODO add your handling code here:
    }//GEN-LAST:event_jLabel73MouseEntered

    private void jLabel73MouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel73MouseExited
        // TODO add your handling code here:
    }//GEN-LAST:event_jLabel73MouseExited

    private void inv_search_tabMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_inv_search_tabMouseClicked
        if (inv == "true") {
            setColor(inv_search_tab);
            resetColor(com_tab);
            resetColor(lap_tab);
            resetColor(home_tab);
            resetColor(user_tab);
            resetColor(search_tab);
            resetColor(viber_tab);
            jTabbedPane2.setSelectedIndex(2);
            Assets.setSelectedIndex(0);
            Fetch1();
            search.requestFocus();
        } else {
            resetColor(inv_search_tab);
        }
    }//GEN-LAST:event_inv_search_tabMouseClicked

    private void inv_search_tabMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_inv_search_tabMouseEntered
        // TODO add your handling code here:
    }//GEN-LAST:event_inv_search_tabMouseEntered

    private void inv_search_tabMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_inv_search_tabMouseExited
        // TODO add your handling code here:
    }//GEN-LAST:event_inv_search_tabMouseExited

    private void inv_search_tabMousePressed(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_inv_search_tabMousePressed
        // TODO add your handling code here:
    }//GEN-LAST:event_inv_search_tabMousePressed

    private void SaveDB6ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_SaveDB6ActionPerformed
        DefaultTableModel model = (DefaultTableModel) user_table.getModel();
        int rowCount = model.getRowCount();
        if (rowCount == 0) {
            JOptionPane.showMessageDialog(null, "You neeed to import files to print!");
        } else {
            MessageFormat header = new MessageFormat("SAP Users");
            MessageFormat footer = new MessageFormat("");
            try {
                user_table.print(JTable.PrintMode.FIT_WIDTH, header, footer);
            } catch (PrinterException e) {
                JOptionPane.showMessageDialog(null, "Unable to Print!");
            }
        }

    }//GEN-LAST:event_SaveDB6ActionPerformed

    private void SaveDB11ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_SaveDB11ActionPerformed
        try {
            DefaultTableModel model = (DefaultTableModel) inv_search_table.getModel();
            if (inv_search_table.getSelectedRowCount() == 0) {
                JOptionPane.showMessageDialog(this, "User not selected");
            } else {
                int row = inv_search_table.getSelectedRow();
                String a = Assets.getSelectedItem().toString();
                String selected = null;

                if (a.equals("Computer")) {
                    pst = con.prepareStatement("DELETE FROM computer_inventory WHERE IP_ADDRESS=?");
                    selected = (String) inv_search_table.getModel().getValueAt(row, 21);
                } else if (a.equals("Laptop")) {
                    pst = con.prepareStatement("DELETE FROM laptop_inventory WHERE asset_id=?");
                    selected = (String) inv_search_table.getModel().getValueAt(row, 1);
                }

                pst.setString(1, selected);
                int k = pst.executeUpdate();
                if (k >= 1) {
                    JOptionPane.showMessageDialog(null, "User deleted");
                    Fetch1();
                } else {
                    JOptionPane.showMessageDialog(null, "User not deleted" + k + selected);
                }
            }
        } catch (SQLException ex) {
            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_SaveDB11ActionPerformed

    private void SaveDB9ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_SaveDB9ActionPerformed
        String a = Assets.getSelectedItem().toString();
        if (a == "Computer") {
            jTabbedPane2.setSelectedIndex(3);
        } else {
            jTabbedPane2.setSelectedIndex(4);
        }

    }//GEN-LAST:event_SaveDB9ActionPerformed

    private void SaveDB10ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_SaveDB10ActionPerformed
        String s = search.getText();

        if (s.isEmpty()) {
            JOptionPane.showMessageDialog(null, "You need to input first");
            DefaultTableModel df = (DefaultTableModel) inv_search_table.getModel();
            df.setRowCount(0);
            Fetch1();
        } else {
            search_inv();
            search.setText("");
            search.requestFocus();
        }
    }//GEN-LAST:event_SaveDB10ActionPerformed

    private void AssetsActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_AssetsActionPerformed

        DefaultTableModel model = (DefaultTableModel) inv_search_table.getModel();
        model.setColumnCount(0);
        settingtable();
        Fetch1();
    }//GEN-LAST:event_AssetsActionPerformed

    private void l_printActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_l_printActionPerformed
        DefaultTableModel model = (DefaultTableModel) laptop_table.getModel();
        int rowCount = model.getRowCount();
        if (rowCount == 0) {
            JOptionPane.showMessageDialog(null, "You neeed to import files to print!");
        } else {
            MessageFormat header = new MessageFormat("Laptop Inventory");
            MessageFormat footer = new MessageFormat("");
            try {
                laptop_table.print(JTable.PrintMode.FIT_WIDTH, header, footer);
            } catch (PrinterException e) {
                JOptionPane.showMessageDialog(null, "Unable to Print!");
            }
        }
    }//GEN-LAST:event_l_printActionPerformed

    private void l_saveActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_l_saveActionPerformed
        DefaultTableModel model = (DefaultTableModel) laptop_table.getModel();
        int rowCount = model.getRowCount();
        int saved = 0;
        try {
            if (rowCount != 0) {
                int rowCount1 = laptop_table.getRowCount();
                int columnCount = laptop_table.getColumnCount();
                for (int row1 = 0; row1 < rowCount1; row1++) {
                    Object dept = laptop_table.getValueAt(row1, 0);
                    Object asst_id = laptop_table.getValueAt(row1, 1);
                    Object asst_desc = laptop_table.getValueAt(row1, 2);
                    Object brand = laptop_table.getValueAt(row1, 3);
                    Object models = laptop_table.getValueAt(row1, 4);
                    Object S_num = laptop_table.getValueAt(row1, 5);
                    Object account = laptop_table.getValueAt(row1, 6);
                    Object date = laptop_table.getValueAt(row1, 7);
                    Object cond = laptop_table.getValueAt(row1, 8);
                    Object stats = laptop_table.getValueAt(row1, 9);
                    Object reco = laptop_table.getValueAt(row1, 10);

                    if (dept == null) {
                        dept = " ";
                    }
                    if (asst_id == null) {
                        asst_id = " ";
                    }
                    if (asst_desc == null) {
                        asst_desc = " ";
                    }
                    if (brand == null) {
                        brand = " ";
                    }
                    if (models == null) {
                        models = " ";
                    }
                    if (S_num == null) {
                        S_num = " ";
                    }
                    if (account == null) {
                        account = " ";
                    }
                    if (date == null) {
                        date = " ";
                    }
                    if (cond == null) {
                        cond = " ";
                    }
                    if (stats == null) {
                        stats = " ";
                    }
                    if (reco == null) {
                        reco = " ";
                    }

                    pst = con.prepareStatement("Insert into Laptop_Inventory(department,Asset_ID,Asset_Description,Brand,Model,Serial_Number,Accountable_to,"
                            + "Warranty_Date,Conditions,Status,Recommendation) values (?,?,?,?,?,?,?,?,?,?,?)");
                    pst.setString(1, dept.toString());
                    pst.setString(2, asst_id.toString());
                    pst.setString(3, asst_desc.toString());
                    pst.setString(4, brand.toString());
                    pst.setString(5, models.toString());
                    pst.setString(6, S_num.toString());
                    pst.setString(7, account.toString());
                    pst.setString(8, date.toString());
                    pst.setString(9, cond.toString());
                    pst.setString(10, stats.toString());
                    pst.setString(11, reco.toString());

                    saved = pst.executeUpdate();
                }
                if (saved == 1) {
                    JOptionPane.showMessageDialog(null, "SAVED!");
                }
            } else {
                JOptionPane.showMessageDialog(null, "You need to import files");
            }
        } catch (SQLException ex) {
            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);

        }
    }//GEN-LAST:event_l_saveActionPerformed

    private void l_exportActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_l_exportActionPerformed
        DefaultTableModel model = (DefaultTableModel) laptop_table.getModel();
        int rowCount = model.getRowCount();
        if (rowCount == 0) {
            JOptionPane.showMessageDialog(null, "You need to import files");
        } else {
            try {
                String defaultCurrentDirectoryPath = "C:\\Users\\Arnie D\\Downloads\\Arnie\\OJT\\IT DEPT";
                JFileChooser jFileChooser = new JFileChooser(defaultCurrentDirectoryPath);
                jFileChooser.showSaveDialog(this);
                File saveFile = jFileChooser.getSelectedFile();
                if (saveFile != null) {
                    saveFile = new File(saveFile.toString() + ".xlsx");
                    if (saveFile.exists()) {
                        int response = JOptionPane.showConfirmDialog(null, //
                                "Do you want to replace the existing file?", //
                                "Confirm", JOptionPane.YES_NO_OPTION, //
                                JOptionPane.QUESTION_MESSAGE);
                        if (response != JOptionPane.YES_OPTION) {

                            jFileChooser.showSaveDialog(this);
                             saveFile = jFileChooser.getSelectedFile();
                            saveFile = new File(saveFile.toString()+ ".xlsx");
                            Workbook wb = new XSSFWorkbook();
                            Sheet sheet = wb.createSheet("Laptop Inventory");
                            Row rowCol = sheet.createRow(0);
                            for (int c = 0; c < laptop_table.getColumnCount(); c++) {
                                Cell cell = rowCol.createCell(c);
                                cell.setCellValue(laptop_table.getColumnName(c));
                            }
                            for (int r = 0; r < laptop_table.getRowCount(); r++) {
                                Row row = sheet.createRow(r);
                                for (int rc = 0; rc < laptop_table.getColumnCount(); rc++) {
                                    Cell cell = row.createCell(rc);
                                    if (laptop_table.getValueAt(r, rc) != null) {
                                        cell.setCellValue(laptop_table.getValueAt(r, rc).toString());
                                    }
                                }
                            }
                            FileOutputStream out = new FileOutputStream(new File(saveFile.toString()));
                            wb.write(out);
                            wb.close();
                            out.close();
                            openFile(saveFile.toString());

                        }
                    } else if(!saveFile.exists()) {
                        Workbook wb = new XSSFWorkbook();
                        Sheet sheet = wb.createSheet("Laptop Inventory");
                        Row rowCol = sheet.createRow(0);
                        for (int c = 0; c < laptop_table.getColumnCount(); c++) {
                            Cell cell = rowCol.createCell(c);
                            cell.setCellValue(laptop_table.getColumnName(c));
                        }
                        for (int r = 0; r < laptop_table.getRowCount(); r++) {
                            Row row = sheet.createRow(r);
                            for (int rc = 0; rc < laptop_table.getColumnCount(); rc++) {
                                Cell cell = row.createCell(rc);
                                if (laptop_table.getValueAt(r, rc) != null) {
                                    cell.setCellValue(laptop_table.getValueAt(r, rc).toString());
                                }
                            }
                        }
                        FileOutputStream out = new FileOutputStream(new File(saveFile.toString()));
                        wb.write(out);
                        wb.close();
                        out.close();
                        openFile(saveFile.toString());
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Error!");
                }

            } catch (FileNotFoundException e) {
                System.out.println(e);
            } catch (IOException io) {
                System.out.println(io);
            }  
        }
    }//GEN-LAST:event_l_exportActionPerformed

    private void l_importActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_l_importActionPerformed
        int RowCount = laptop_table.getRowCount();
        if (RowCount != 0) {
            DefaultTableModel df = (DefaultTableModel) laptop_table.getModel();
            df.setRowCount(0);
            RowCount = 0;
            if (RowCount == 0) {
                Import_LaptopInventory();
            }
        } else {
            Import_LaptopInventory();
        }
    }//GEN-LAST:event_l_importActionPerformed

    private void c_printActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_c_printActionPerformed
        DefaultTableModel model = (DefaultTableModel) computer_table.getModel();
        int rowCount = model.getRowCount();
        if (rowCount == 0) {
            JOptionPane.showMessageDialog(null, "You neeed to import files to print!");
        } else {
            MessageFormat header = new MessageFormat("Computer Inventory");
            MessageFormat footer = new MessageFormat("");
            try {
                computer_table.print(JTable.PrintMode.FIT_WIDTH, header, footer);
            } catch (PrinterException e) {
                JOptionPane.showMessageDialog(null, "Unable to Print!");
            }
        }
    }//GEN-LAST:event_c_printActionPerformed

    private void c_saveActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_c_saveActionPerformed
        DefaultTableModel model = (DefaultTableModel) computer_table.getModel();
        int rowCount = model.getRowCount();
        int saved = 0;
        try {
            if (rowCount != 0) {
                int rowCount1 = computer_table.getRowCount();
                int columnCount = computer_table.getColumnCount();
                for (int row1 = 0; row1 < rowCount1; row1++) {
                    Object name = computer_table.getValueAt(row1, 0);
                    Object dept = computer_table.getValueAt(row1, 1);
                    Object mbrand = computer_table.getValueAt(row1, 2);
                    Object mab = computer_table.getValueAt(row1, 3);
                    Object mak = computer_table.getValueAt(row1, 4);
                    Object mbm = computer_table.getValueAt(row1, 5);
                    Object msn = computer_table.getValueAt(row1, 6);
                    Object pb = computer_table.getValueAt(row1, 7);
                    Object psn = computer_table.getValueAt(row1, 8);
                    Object hdb = computer_table.getValueAt(row1, 9);
                    Object hds = computer_table.getValueAt(row1, 10);
                    Object hdsn = computer_table.getValueAt(row1, 11);
                    Object mb = computer_table.getValueAt(row1, 12);
                    Object ms = computer_table.getValueAt(row1, 13);
                    Object mmrysn = computer_table.getValueAt(row1, 14);
                    Object gc = computer_table.getValueAt(row1, 15);
                    Object sn = computer_table.getValueAt(row1, 16);
                    Object p = computer_table.getValueAt(row1, 17);
                    Object ps = computer_table.getValueAt(row1, 18);
                    Object ola = computer_table.getValueAt(row1, 19);
                    Object wla = computer_table.getValueAt(row1, 20);
                    Object ia = computer_table.getValueAt(row1, 21);
                    Object yb = computer_table.getValueAt(row1, 22);
                    Object fb = computer_table.getValueAt(row1, 23);
                    Object usb = computer_table.getValueAt(row1, 24);
                    Object d = computer_table.getValueAt(row1, 25);
                    Object w = computer_table.getValueAt(row1, 26);
                    Object h = computer_table.getValueAt(row1, 27);
                    Object wedp = computer_table.getValueAt(row1, 28);

                    if (name == null) {
                        name = " ";
                    }
                    if (dept == null) {
                        dept = " ";
                    }
                    if (mbrand == null) {
                        name = " ";
                    }
                    if (mab == null) {
                        mab = " ";
                    }
                    if (mak == null) {
                        mak = " ";
                    }
                    if (mbm == null) {
                        mbm = " ";
                    }
                    if (msn == null) {
                        msn = " ";
                    }
                    if (pb == null) {
                        pb = " ";
                    }
                    if (psn == null) {
                        psn = " ";
                    }
                    if (hdb == null) {
                        hdb = " ";
                    }
                    if (hds == null) {
                        hds = " ";
                    }
                    if (hdsn == null) {
                        hdsn = " ";
                    }
                    if (mb == null) {
                        mb = " ";
                    }
                    if (ms == null) {
                        ms = " ";
                    }
                    if (mmrysn == null) {
                        mmrysn = " ";
                    }
                    if (gc == null) {
                        gc = " ";
                    }
                    if (sn == null) {
                        sn = " ";
                    }
                    if (p == null) {
                        p = " ";
                    }
                    if (ps == null) {
                        ps = " ";
                    }
                    if (ola == null) {
                        ola = " ";
                    }
                    if (wla == null) {
                        wla = " ";
                    }
                    if (ia == null) {
                        ia = " ";
                    }
                    if (yb == null) {
                        yb = " ";
                    }
                    if (fb == null) {
                        fb = " ";
                    }
                    if (usb == null) {
                        usb = " ";
                    }
                    if (d == null) {
                        d = " ";
                    }
                    if (w == null) {
                        w = " ";
                    }
                    if (h == null) {
                        h = " ";
                    }
                    if (wedp == null) {
                        wedp = " ";
                    }

                    pst = con.prepareStatement("Insert into Computer_Inventory(NAME,DEPARTMENT,MONITOR_BRAND,MONITOR_ASSET_BRAND,MOUSE_AND_KEYBOARD,MOTHERBOARD_BRAND_MODEL,MOTHERBOARD_SERIAL_NO,"
                            + "POWERSUPPLY_BRAND,POWERSUPPLY_SERIAL_NO,HARD_DRIVE_BRAND,HARD_DRIVE_SIZE,HARD_DRIVE_SERIAL_NO,MEMORY_BRAND, MEMORY_SIZE,MEMORY_SERIAL_NO, GRAPHIC_CARDS, SERIAL_NUMBER,PROCESSOR,"
                            + " PROCESSOR_SPECS,OFFICE_LICENSE_ACTIVATED,WINDOWS_LICENSE_ACTIVATED,IP_ADDRESS,YOUTUBE_BLOCKED,FB_BLOCKED,USB_ENABLED,DOMAIN,WEBCAM,HEADSET,WARRANTY_END_DATE_PROCESSOR)"
                            + " values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
                    pst.setString(1, name.toString());
                    pst.setString(2, dept.toString());
                    pst.setString(3, mbrand.toString());
                    pst.setString(4, mab.toString());
                    pst.setString(5, mak.toString());
                    pst.setString(6, mbm.toString());
                    pst.setString(7, msn.toString());
                    pst.setString(8, pb.toString());
                    pst.setString(9, psn.toString());
                    pst.setString(10, hdb.toString());
                    pst.setString(11, hds.toString());
                    pst.setString(12, hdsn.toString());
                    pst.setString(13, mb.toString());
                    pst.setString(14, ms.toString());
                    pst.setString(15, mmrysn.toString());
                    pst.setString(16, gc.toString());
                    pst.setString(17, sn.toString());
                    pst.setString(18, p.toString());
                    pst.setString(19, ps.toString());
                    pst.setString(20, ola.toString());
                    pst.setString(21, wla.toString());
                    pst.setString(22, ia.toString());
                    pst.setString(23, yb.toString());
                    pst.setString(24, fb.toString());
                    pst.setString(25, usb.toString());
                    pst.setString(26, d.toString());
                    pst.setString(27, w.toString());
                    pst.setString(28, h.toString());
                    pst.setString(29, wedp.toString());

                    saved = pst.executeUpdate();
                }
                if (saved == 1) {
                    JOptionPane.showMessageDialog(null, "SAVED!");
                }
            } else {
                JOptionPane.showMessageDialog(null, "You need to import files");
            }
        } catch (SQLException ex) {
            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);

        }
    }//GEN-LAST:event_c_saveActionPerformed

    private void c_exportActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_c_exportActionPerformed

          DefaultTableModel model = (DefaultTableModel) computer_table.getModel();
        int rowCount = model.getRowCount();
        if (rowCount == 0) {
            JOptionPane.showMessageDialog(null, "You need to import files");
        } else {
            try {
                String defaultCurrentDirectoryPath = "C:\\Users\\Arnie D\\Downloads\\Arnie\\OJT\\IT DEPT";
                JFileChooser jFileChooser = new JFileChooser(defaultCurrentDirectoryPath);
                jFileChooser.showSaveDialog(this);
                File saveFile = jFileChooser.getSelectedFile();
                if (saveFile != null) {
                    saveFile = new File(saveFile.toString() + ".xlsx");
                    if (saveFile.exists()) {
                        int response = JOptionPane.showConfirmDialog(null, //
                                "Do you want to replace the existing file?", //
                                "Confirm", JOptionPane.YES_NO_OPTION, //
                                JOptionPane.QUESTION_MESSAGE);
                        if (response != JOptionPane.YES_OPTION) {

                            jFileChooser.showSaveDialog(this);
                             saveFile = jFileChooser.getSelectedFile();
                            saveFile = new File(saveFile.toString()+ ".xlsx");
                            Workbook wb = new XSSFWorkbook();
                            Sheet sheet = wb.createSheet("Computer Inventory");
                            Row rowCol = sheet.createRow(0);
                            for (int c = 0; c < computer_table.getColumnCount(); c++) {
                                Cell cell = rowCol.createCell(c);
                                cell.setCellValue(computer_table.getColumnName(c));
                            }
                            for (int r = 0; r < computer_table.getRowCount(); r++) {
                                Row row = sheet.createRow(r);
                                for (int rc = 0; rc < computer_table.getColumnCount(); rc++) {
                                    Cell cell = row.createCell(rc);
                                    if (computer_table.getValueAt(r, rc) != null) {
                                        cell.setCellValue(computer_table.getValueAt(r, rc).toString());
                                    }
                                }
                            }
                            FileOutputStream out = new FileOutputStream(new File(saveFile.toString()));
                            wb.write(out);
                            wb.close();
                            out.close();
                            openFile(saveFile.toString());

                        }
                    } else if(!saveFile.exists()) {
                        Workbook wb = new XSSFWorkbook();
                        Sheet sheet = wb.createSheet("EMAIL LIST ACCOUNTS");
                        Row rowCol = sheet.createRow(0);
                        for (int c = 0; c < computer_table.getColumnCount(); c++) {
                            Cell cell = rowCol.createCell(c);
                            cell.setCellValue(computer_table.getColumnName(c));
                        }
                        for (int r = 0; r < computer_table.getRowCount(); r++) {
                            Row row = sheet.createRow(r);
                            for (int rc = 0; rc < computer_table.getColumnCount(); rc++) {
                                Cell cell = row.createCell(rc);
                                if (computer_table.getValueAt(r, rc) != null) {
                                    cell.setCellValue(computer_table.getValueAt(r, rc).toString());
                                }
                            }
                        }
                        FileOutputStream out = new FileOutputStream(new File(saveFile.toString()));
                        wb.write(out);
                        wb.close();
                        out.close();
                        openFile(saveFile.toString());
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Error!");
                }

            } catch (FileNotFoundException e) {
                System.out.println(e);
            } catch (IOException io) {
                System.out.println(io);
            }  
        }
    }//GEN-LAST:event_c_exportActionPerformed

    private void c_importActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_c_importActionPerformed
        int RowCount = computer_table.getRowCount();
        if (RowCount != 0) {
            DefaultTableModel df = (DefaultTableModel) computer_table.getModel();
            df.setRowCount(0);
            RowCount = 0;
            if (RowCount == 0) {
                Import_Inventory();
            }
        } else {
            Import_Inventory();
        }
    }//GEN-LAST:event_c_importActionPerformed

    private void SaveDB12ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_SaveDB12ActionPerformed
        try {
            String department = dept1.getText();
            String asst = id.getText();
            String description = desc.getText();
            String brand = brnd.getText();
            String model = mdl.getText();
            String serial = srl.getText();
            String accountable = acct.getText();
            java.util.Date warrant = c_wedp.getDate();
            String cond = condi.getText();
            String stats = status.getText();
            String recommendation = reco.getText();
            String warranty = null;
           
            if(department.equals("")||asst.equals("")||description.equals("")||brand.equals("")||model.equals("")
                    ||serial.equals("")||accountable.equals("")||warrant==null ||cond.equals("")||stats.equals("")||recommendation.equals("")){
                
                JOptionPane.showMessageDialog(this,"Please put all the necessary information.");
            }
            else{
               SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd", Locale.ENGLISH);
                warranty = dateFormat.format(warrant);
                
            pst = con.prepareStatement("Insert into laptop_inventory(department, asset_id,asset_description,brand,model,serial_number,accountable_to,"
                    + "warranty_date,conditions,status,recommendation) values (?,?,?,?,?,?,?,?,?,?,?)");
            pst.setString(1, department);
            pst.setString(2, asst);
            pst.setString(3, description);
            pst.setString(4, brand);
            pst.setString(5, model);
            pst.setString(6, serial);
            pst.setString(7, accountable);
            pst.setString(8, warranty);
            pst.setString(9, cond);
            pst.setString(10, stats);
            pst.setString(11, recommendation);

            int k = pst.executeUpdate();

            if (k == 1) {
                JOptionPane.showMessageDialog(this, "Inventory updated");
                Fetch1();
                dept1.setText("");
                id.setText("");
                desc.setText("");
                brnd.setText("");
                mdl.setText("");
                srl.setText("");
                acct.setText("");
                c_wedp.setCalendar(null);
                condi.setText("");
                status.setText("");
                reco.setText("");
                jTabbedPane2.setSelectedIndex(2);

            } else {
                JOptionPane.showMessageDialog(this, "Inventory not updated");
            }
            }
        } catch (SQLException ex) {
            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);
        }
        

    }//GEN-LAST:event_SaveDB12ActionPerformed

    private void l_print1ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_l_print1ActionPerformed
        DefaultTableModel model = (DefaultTableModel) inv_search_table.getModel();
        int rowCount = model.getRowCount();
        String a = Assets.getSelectedItem().toString();
        MessageFormat header;
        MessageFormat footer;
        if (rowCount == 0) {
            JOptionPane.showMessageDialog(null, "You neeed to import files to print!");
        } else {
            if (a == "Computer") {
                header = new MessageFormat("Computer Inventory");
                footer = new MessageFormat("");
            } else {
                header = new MessageFormat("Laptop Inventory");
                footer = new MessageFormat("");
            }

            try {
                inv_search_table.print(JTable.PrintMode.FIT_WIDTH, header, footer);
            } catch (PrinterException e) {
                JOptionPane.showMessageDialog(null, "Unable to Print!");
            }
        }
    }//GEN-LAST:event_l_print1ActionPerformed

    private void SaveDB13ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_SaveDB13ActionPerformed
        try {
            String name = c_n.getText();
            String dept = c_dept.getText();
            String mbrand = c_mb.getText();
            String mab = c_mab.getText();
            String mak = c_mak.getText();
            String mbm = c_mbm.getText();
            String msn = c_msn.getText();
            String pb = c_psb.getText();
            String psn = c_pssn.getText();
            String hdb = c_hdb.getText();
            String hds = c_hds.getText();
            String hdsn = c_hdsn.getText();
            String mb = c_mb.getText();
            String ms = c_ms.getText();
            String mmrysn = c_mmrysn.getText();
            String gc = c_gc.getText();
            String sn = c_sn.getText();
            String p = c_p.getText();
            String ps = c_n.getText();
            String ola = c_oal.getSelectedItem().toString();
            String wla = c_wla.getSelectedItem().toString();
            String ia = c_ip.getText();
            String yb = c_yt.getSelectedItem().toString();
            String fb = c_fb.getSelectedItem().toString();
            String usb = c_usb.getSelectedItem().toString();
            String d = c_d.getSelectedItem().toString();
            String w = c_w.getSelectedItem().toString();
            String h = c_h.getSelectedItem().toString();

            java.util.Date warrant = c_wedp.getDate();
            String warranty = null;
            
            if(name.equals("")||  dept.equals("") ||mbrand.equals("") || mab.equals("")|| mak.equals("") 
                    || mbm.equals("") || msn.equals("") || pb.equals("")||psn.equals("")||hdb.equals("") ||hds.equals("")
                    || hdsn.equals("")|| mb.equals("") ||ms.equals("")||mmrysn.equals("")||gc.equals("")||sn.equals("")
                    || p.equals("")||ps.equals("")||ola.equals("")||wla.equals("")||ia.equals("")||yb.equals("")
                    ||fb.equals("")||usb.equals("")||d.equals("")||w.equals("")||h.equals("") || warrant == null){
                JOptionPane.showMessageDialog(this,"Please put all the necessary information.");
            }
            else{
                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd", Locale.ENGLISH);
                warranty = dateFormat.format(warrant);
            pst = con.prepareStatement("Insert into Computer_Inventory(NAME,DEPARTMENT,MONITOR_BRAND,MONITOR_ASSET_BRAND,MOUSE_AND_KEYBOARD,MOTHERBOARD_BRAND_MODEL,MOTHERBOARD_SERIAL_NO,"
                    + "POWERSUPPLY_BRAND,POWERSUPPLY_SERIAL_NO,HARD_DRIVE_BRAND,HARD_DRIVE_SIZE,HARD_DRIVE_SERIAL_NO,MEMORY_BRAND, MEMORY_SIZE,MEMORY_SERIAL_NO, GRAPHIC_CARDS, SERIAL_NUMBER,PROCESSOR,"
                    + " PROCESSOR_SPECS,OFFICE_LICENSE_ACTIVATED,WINDOWS_LICENSE_ACTIVATED,IP_ADDRESS,YOUTUBE_BLOCKED,FB_BLOCKED,USB_ENABLED,DOMAIN,WEBCAM,HEADSET,WARRANTY_END_DATE_PROCESSOR)"
                    + " values (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)");
            pst.setString(1, name);
            pst.setString(2, dept);
            pst.setString(3, mbrand);
            pst.setString(4, mab);
            pst.setString(5, mak);
            pst.setString(6, mbm);
            pst.setString(7, msn);
            pst.setString(8, pb);
            pst.setString(9, psn);
            pst.setString(10, hdb);
            pst.setString(11, hds);
            pst.setString(12, hdsn);
            pst.setString(13, mb);
            pst.setString(14, ms);
            pst.setString(15, mmrysn);
            pst.setString(16, gc);
            pst.setString(17, sn);
            pst.setString(18, p);
            pst.setString(19, ps);
            pst.setString(20, ola);
            pst.setString(21, wla);
            pst.setString(22, ia);
            pst.setString(23, yb);
            pst.setString(24, fb);
            pst.setString(25, usb);
            pst.setString(26, d);
            pst.setString(27, w);
            pst.setString(28, h);
            pst.setString(29, warranty);
            
            int k = pst.executeUpdate();
            
            if (k == 1) {
                JOptionPane.showMessageDialog(this, "Inventory updated");
                Fetch1();
                jTabbedPane2.setSelectedIndex(2);

            } else {
                JOptionPane.showMessageDialog(this, "Inventory not updated");
            }
            }
        }catch (SQLException ex) {
            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);
        }
        
    }//GEN-LAST:event_SaveDB13ActionPerformed

    private void c_wlaActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_c_wlaActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_c_wlaActionPerformed

    private void c_oalActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_c_oalActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_c_oalActionPerformed

    private void c_ytActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_c_ytActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_c_ytActionPerformed

    private void c_fbActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_c_fbActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_c_fbActionPerformed

    private void c_usbActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_c_usbActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_c_usbActionPerformed

    private void c_dActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_c_dActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_c_dActionPerformed

    private void c_hActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_c_hActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_c_hActionPerformed

    private void c_wActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_c_wActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_c_wActionPerformed

    private void jLabel47MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel47MouseClicked
        jTabbedPane1.setSelectedIndex(3);
        jTabbedPane2.setSelectedIndex(0);
        setColor(inventory1);
        setColor(com_tab);
        resetColor(email_tab);
        resetColor(home_tab);
        resetColor(user_tab);
        resetColor(search_tab);
        resetColor(viber_tab);
        setTable(computer_table);
        setCInventory(computer_table);

        jLabel62.setForeground(Color.white);
        jLabel48.setForeground(Color.white);
        jLabel66.setForeground(Color.white);
        jLabel55.setVisible(true);
        jLabel63.setVisible(true);
        jLabel73.setVisible(true);
        Icon computer = jLabel55.getIcon();
        ImageIcon iconc = (ImageIcon) computer;
        Image imagec = iconc.getImage().getScaledInstance(jLabel55.getWidth(), jLabel55.getHeight(), Image.SCALE_SMOOTH);
        jLabel55.setIcon(new ImageIcon(imagec));

        Icon laptop = jLabel63.getIcon();
        ImageIcon iconl = (ImageIcon) laptop;
        Image imagel = iconl.getImage().getScaledInstance(jLabel63.getWidth(), jLabel63.getHeight(), Image.SCALE_SMOOTH);
        jLabel63.setIcon(new ImageIcon(imagel));
        inv = "true";

        Icon s = jLabel73.getIcon();
        ImageIcon icons = (ImageIcon) s;
        Image images = icons.getImage().getScaledInstance(jLabel73.getWidth(), jLabel73.getHeight(), Image.SCALE_SMOOTH);
        jLabel73.setIcon(new ImageIcon(images));
        inv = "true";
    }//GEN-LAST:event_jLabel47MouseClicked

    private void jLabel14MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel14MouseClicked
        jTabbedPane1.setSelectedIndex(2);
        resetColor(email_tab);
        setColor(user_tab);
        resetColor(search_tab);
        resetColor(home_tab);
        resetColor(inventory1);
        resetColor(com_tab);
        resetColor(lap_tab);
        resetColor(inv_search_tab);
        resetColor(viber_tab);
        setTable(user_table);
        jLabel62.setForeground(new Color(228, 57, 39));
        jLabel48.setForeground(new Color(228, 57, 39));
        jLabel66.setForeground(new Color(228, 57, 39));
        jLabel55.setVisible(false);
        jLabel63.setVisible(false);
        jLabel73.setVisible(false);
        inv = "false";

    }//GEN-LAST:event_jLabel14MouseClicked

    private void jLabel15MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel15MouseClicked
        jTabbedPane1.setSelectedIndex(2);
        resetColor(email_tab);
        setColor(user_tab);
        resetColor(search_tab);
        resetColor(home_tab);
        resetColor(inventory1);
        resetColor(com_tab);
        resetColor(lap_tab);
        resetColor(inv_search_tab);
        resetColor(viber_tab);
        setTable(user_table);
        jLabel62.setForeground(new Color(228, 57, 39));
        jLabel48.setForeground(new Color(228, 57, 39));
        jLabel66.setForeground(new Color(228, 57, 39));
        jLabel55.setVisible(false);
        jLabel63.setVisible(false);
        jLabel73.setVisible(false);
        inv = "false";

    }//GEN-LAST:event_jLabel15MouseClicked

    private void jLabel39MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel39MouseClicked
        user_code.requestFocus();
        user_search.requestFocus();
        jTabbedPane1.setSelectedIndex(1);
        resetColor(email_tab);
        setTable(search_table);
        setColor(search_tab);
        resetColor(home_tab);
        resetColor(user_tab);
        resetColor(inventory1);
        resetColor(com_tab);
        resetColor(lap_tab);
        resetColor(inv_search_tab);
        resetColor(viber_tab);
        jLabel62.setForeground(new Color(228, 57, 39));
        jLabel48.setForeground(new Color(228, 57, 39));
        jLabel66.setForeground(new Color(228, 57, 39));
        jLabel55.setVisible(false);
        jLabel63.setVisible(false);
        jLabel73.setVisible(false);

        com_tab.setEnabled(false);
        inv = "false";
        isEmpty();
    }//GEN-LAST:event_jLabel39MouseClicked

    private void l_print2ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_l_print2ActionPerformed
        DefaultTableModel model = (DefaultTableModel) search_table.getModel();
        int rowCount = model.getRowCount();
        if (rowCount == 0) {
            JOptionPane.showMessageDialog(null, "You neeed to import files to print!");
        } else {
            MessageFormat header = new MessageFormat("SAP Users");
            MessageFormat footer = new MessageFormat("");
            try {
                search_table.print(JTable.PrintMode.FIT_WIDTH, header, footer);
            } catch (PrinterException e) {
                JOptionPane.showMessageDialog(null, "Unable to Print!");
            }
        }
    }//GEN-LAST:event_l_print2ActionPerformed

    private void jLabel89MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel89MouseClicked
        jTabbedPane1.setSelectedIndex(4);
        resetColor(email_tab);
        resetColor(user_tab);
        resetColor(search_tab);
        resetColor(home_tab);
        resetColor(inventory1);
        resetColor(com_tab);
        resetColor(lap_tab);
        resetColor(inv_search_tab);
        setColor(viber_tab);
        ViberTable(viber_table);
        Fetch_Viber();
        jLabel62.setForeground(new Color(228, 57, 39));
        jLabel48.setForeground(new Color(228, 57, 39));
        jLabel66.setForeground(new Color(228, 57, 39));
        jLabel55.setVisible(false);
        jLabel63.setVisible(false);
        jLabel73.setVisible(false);
        inv = "false";
    }//GEN-LAST:event_jLabel89MouseClicked

    private void jLabel90MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel90MouseClicked
        jTabbedPane1.setSelectedIndex(4);
        resetColor(email_tab);
        resetColor(user_tab);
        resetColor(search_tab);
        resetColor(home_tab);
        resetColor(inventory1);
        resetColor(com_tab);
        resetColor(lap_tab);
        resetColor(inv_search_tab);
        setColor(viber_tab);
        ViberTable(viber_table);
        Fetch_Viber();
        jLabel62.setForeground(new Color(228, 57, 39));
        jLabel48.setForeground(new Color(228, 57, 39));
        jLabel66.setForeground(new Color(228, 57, 39));
        jLabel55.setVisible(false);
        jLabel63.setVisible(false);
        jLabel73.setVisible(false);
        inv = "false";
    }//GEN-LAST:event_jLabel90MouseClicked

    private void viber_tabMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_viber_tabMouseClicked
        jTabbedPane1.setSelectedIndex(4);
        resetColor(email_tab);
        resetColor(user_tab);
        resetColor(search_tab);
        resetColor(home_tab);
        resetColor(inventory1);
        resetColor(com_tab);
        resetColor(lap_tab);
        resetColor(inv_search_tab);
        setColor(viber_tab);
        ViberTable(viber_table);
        Fetch_Viber();
        jLabel62.setForeground(new Color(228, 57, 39));
        jLabel48.setForeground(new Color(228, 57, 39));
        jLabel66.setForeground(new Color(228, 57, 39));
        jLabel55.setVisible(false);
        jLabel63.setVisible(false);
        jLabel73.setVisible(false);
        inv = "false";
    }//GEN-LAST:event_viber_tabMouseClicked

    private void viber_tabMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_viber_tabMouseEntered
        // TODO add your handling code here:
    }//GEN-LAST:event_viber_tabMouseEntered

    private void viber_tabMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_viber_tabMouseExited
        // TODO add your handling code here:
    }//GEN-LAST:event_viber_tabMouseExited

    private void viber_tabMousePressed(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_viber_tabMousePressed
        // TODO add your handling code here:
    }//GEN-LAST:event_viber_tabMousePressed

    private void jLabel47MouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel47MouseEntered
        // TODO add your handling code here:
    }//GEN-LAST:event_jLabel47MouseEntered

    private void SaveDB14ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_SaveDB14ActionPerformed
        DefaultTableModel model = (DefaultTableModel) viber_table.getModel();
        int rowCount = model.getRowCount();
        if (rowCount == 0) {
            JOptionPane.showMessageDialog(null, "You neeed to import files to print!");
            Fetch_Viber();
        } else {
            MessageFormat header = new MessageFormat("VIBER ACCOUNTS");
            MessageFormat footer = new MessageFormat("");
            try {
                viber_table.print(JTable.PrintMode.FIT_WIDTH, header, footer);
            } catch (PrinterException e) {
                JOptionPane.showMessageDialog(null, "Unable to Print!");
            }
        }
    }//GEN-LAST:event_SaveDB14ActionPerformed

    private void Viber_addActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_Viber_addActionPerformed
        int saved = 0;
        try {
            String client = client_name.getText();
            String num = mobile_number.getText();
            String dept = department.getText();
            String dvt = device_type.getSelectedItem().toString();
            if (client.equals("") || (num.equals("")) || (dept.equals("")) || dvt.equals("Device Type:")) {
                JOptionPane.showMessageDialog(null, "Please input all the necessary information");
                client_name.requestFocus();
            } else {
                String dup = "Select * from viber_accounts where mobile_number='" + num + "'";
                pst = con.prepareStatement(dup);
                rs = pst.executeQuery(dup);

                if (rs.next()) {

                    JOptionPane.showMessageDialog(null, "Mobile number has been already used.");
                    mobile_number.requestFocus();

                } else {
                    String dup1 = "Select * from viber_accounts where client_name='" + client + "'";
                    pst = con.prepareStatement(dup1);
                    rs = pst.executeQuery(dup1);
                    if (rs.next()) {
                        JOptionPane.showMessageDialog(null, "Client name has been already used.");
                        client_name.requestFocus();
                    } else {
                        pst = con.prepareStatement("Insert into viber_accounts( client_name,mobile_number,department,device_type) values (?,?,?,?)");
                        pst.setString(1, client);
                        pst.setString(2, num);
                        pst.setString(3, dept);
                        pst.setString(4, dvt);

                        saved = pst.executeUpdate();
                    }

                }
                if (saved == 1) {
                    JOptionPane.showMessageDialog(null, "SAVED!");
                    Fetch_Viber();
                    client_name.setText("");
                    mobile_number.setText("");
                    department.setText("");
                    device_type.setSelectedIndex(0);

                }
            }

        } catch (SQLException ex) {
            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);

        }
    }//GEN-LAST:event_Viber_addActionPerformed

    private void Export_ViberActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_Export_ViberActionPerformed
        DefaultTableModel model = (DefaultTableModel) viber_table.getModel();
        int rowCount = model.getRowCount();
        if (rowCount == 0) {
            JOptionPane.showMessageDialog(null, "You need to import files");
        } else {
            try {
                String defaultCurrentDirectoryPath = "C:\\Users\\Arnie D\\Downloads\\Arnie\\OJT\\IT DEPT";
                JFileChooser jFileChooser = new JFileChooser(defaultCurrentDirectoryPath);
                jFileChooser.showSaveDialog(this);
                File saveFile = jFileChooser.getSelectedFile();
                if (saveFile != null) {
                    saveFile = new File(saveFile.toString() + ".xlsx");
                    if (saveFile.exists()) {
                        int response = JOptionPane.showConfirmDialog(null, //
                                "Do you want to replace the existing file?", //
                                "Confirm", JOptionPane.YES_NO_OPTION, //
                                JOptionPane.QUESTION_MESSAGE);
                        if (response != JOptionPane.YES_OPTION) {

                            jFileChooser.showSaveDialog(this);
                             saveFile = jFileChooser.getSelectedFile();
                            saveFile = new File(saveFile.toString()+ ".xlsx");
                            Workbook wb = new XSSFWorkbook();
                            Sheet sheet = wb.createSheet("VIBER ACCOUNTS");
                            Row rowCol = sheet.createRow(0);
                            for (int c = 0; c < viber_table.getColumnCount(); c++) {
                                Cell cell = rowCol.createCell(c);
                                cell.setCellValue(viber_table.getColumnName(c));
                            }
                            for (int r = 0; r < viber_table.getRowCount(); r++) {
                                Row row = sheet.createRow(r);
                                for (int rc = 0; rc < viber_table.getColumnCount(); rc++) {
                                    Cell cell = row.createCell(rc);
                                    if (viber_table.getValueAt(r, rc) != null) {
                                        cell.setCellValue(viber_table.getValueAt(r, rc).toString());
                                    }
                                }
                            }
                            FileOutputStream out = new FileOutputStream(new File(saveFile.toString()));
                            wb.write(out);
                            wb.close();
                            out.close();
                            openFile(saveFile.toString());

                        }
                    } else if(!saveFile.exists()) {
                        Workbook wb = new XSSFWorkbook();
                        Sheet sheet = wb.createSheet("VIBER ACCOUNTS");
                        Row rowCol = sheet.createRow(0);
                        for (int c = 0; c < viber_table.getColumnCount(); c++) {
                            Cell cell = rowCol.createCell(c);
                            cell.setCellValue(viber_table.getColumnName(c));
                        }
                        for (int r = 0; r < viber_table.getRowCount(); r++) {
                            Row row = sheet.createRow(r);
                            for (int rc = 0; rc < viber_table.getColumnCount(); rc++) {
                                Cell cell = row.createCell(rc);
                                if (viber_table.getValueAt(r, rc) != null) {
                                    cell.setCellValue(viber_table.getValueAt(r, rc).toString());
                                }
                            }
                        }
                        FileOutputStream out = new FileOutputStream(new File(saveFile.toString()));
                        wb.write(out);
                        wb.close();
                        out.close();
                        openFile(saveFile.toString());
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Error!");
                }

            } catch (FileNotFoundException e) {
                System.out.println(e);
            } catch (IOException io) {
                System.out.println(io);
            }  
        }
    }//GEN-LAST:event_Export_ViberActionPerformed

    private void DeleteActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_DeleteActionPerformed
        try {
            DefaultTableModel model = (DefaultTableModel) viber_table.getModel();
            if (viber_table.getSelectedRowCount() == 0) {
                JOptionPane.showMessageDialog(this, "Client not selected");
            } else {
                int row = viber_table.getSelectedRow();
                Object client = viber_table.getModel().getValueAt(row, 0);
                Object number = viber_table.getModel().getValueAt(row, 1);
                Object dept = viber_table.getModel().getValueAt(row, 2);
                Object device = viber_table.getModel().getValueAt(row, 3);
                pst = con.prepareStatement("DELETE FROM viber_accounts WHERE client_name=? and mobile_number=? and department=? and device_type=?");
                pst.setString(1, client.toString());
                pst.setString(2, number.toString());
                pst.setString(3, dept.toString());
                pst.setString(4, device.toString());
                int k = pst.executeUpdate();
                if (k == 1) {
                    JOptionPane.showMessageDialog(this, "User deleted");
                    Fetch_Viber();
                } else {
                    JOptionPane.showMessageDialog(this, "User not deleted");
                }
            }
        } catch (SQLException ex) {
            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_DeleteActionPerformed

    private void Viber_editActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_Viber_editActionPerformed
        try {
            String cn = client_name.getText();
            String mn = mobile_number.getText();
            String dept = department.getText();
            String dvt = device_type.getSelectedItem().toString();
            int row = viber_table.getSelectedRow();
            Object number = viber_table.getModel().getValueAt(row, 1);

            if (mn.equals(number)) {
                pst = con.prepareStatement("UPDATE viber_accounts SET client_name =?, mobile_number=?, department=?, device_type=? WHERE mobile_number=? ");
                pst.setString(1, cn);
                pst.setString(2, mn);
                pst.setString(3, dept);
                pst.setString(4, dvt);
                pst.setString(5, number.toString());
                int k = pst.executeUpdate();

                if (k == 1) {
                    JOptionPane.showMessageDialog(this, "Viber account updated");
                    Fetch_Viber();
                } else {
                    JOptionPane.showMessageDialog(this, "Viber account not updated");
                }
            } else {
                String duplicate = "SELECT * FROM viber_accounts WHERE mobile_number = '" + mn + "'";
                pst = con.prepareStatement(duplicate);
                rs = pst.executeQuery(duplicate);
                if (rs.next()) {
                    JOptionPane.showMessageDialog(this, "Mobile number already used");
                    mobile_number.requestFocus();
                } else {
                    pst = con.prepareStatement("UPDATE viber_accounts SET client_name =?, mobile_number=?, department=?, device_type=? WHERE mobile_number=? ");
                    pst.setString(1, cn);
                    pst.setString(2, mn);
                    pst.setString(3, dept);
                    pst.setString(4, dvt);
                    pst.setString(5, number.toString());
                    int k = pst.executeUpdate();

                    if (k == 1) {
                        JOptionPane.showMessageDialog(this, "Viber account updated");
                        Fetch_Viber();
                    } else {
                        JOptionPane.showMessageDialog(this, "Viber account not updated");
                    }
                }
            }

        } catch (SQLException ex) {
            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_Viber_editActionPerformed

    private void Viber_ImportActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_Viber_ImportActionPerformed
        int RowCount = viber_table.getRowCount();
        if (RowCount != 0) {
            DefaultTableModel df = (DefaultTableModel) viber_table.getModel();
            df.setRowCount(0);
            RowCount = 0;
            if (RowCount == 0) {
                Import_Viber();
            }
        } else {
            Import_Viber();
        }
    }//GEN-LAST:event_Viber_ImportActionPerformed

    private void SaveDB20ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_SaveDB20ActionPerformed
        String v = viber_search.getText();

        if (v.isEmpty()) {
            JOptionPane.showMessageDialog(null, "You need to input first");
            DefaultTableModel df = (DefaultTableModel) viber_table.getModel();
            df.setRowCount(0);
            Fetch_Viber();
        } else {
            search_viber();
            search.setText("");
        }
    }//GEN-LAST:event_SaveDB20ActionPerformed

    private void user_editActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_user_editActionPerformed
        try {
            String code = user_code.getText();
            String nm = user_name.getText();
            String dp = user_department.getText();
            String act = user_activity.getText();
            String db1 = user_database.getText();
            String lt = LType.getSelectedItem().toString();
            String r = user_remark.getText();
            int row = search_table.getSelectedRow();
            Object user_code1 = search_table.getModel().getValueAt(row, 3);

            if (code.equals(user_code1)) {
                pst = con.prepareStatement("UPDATE user SET department=?,activity=?,db=?,user_code=?,name=?,license=?,remarks=? where user_code=?");
                pst.setString(1, dp);
                pst.setString(2, act);
                pst.setString(3, db1);
                pst.setString(4, code);
                pst.setString(5, nm);
                pst.setString(6, lt);
                pst.setString(7, r);
                pst.setString(8, user_code1.toString());
                int k = pst.executeUpdate();

                if (k == 1) {
                    JOptionPane.showMessageDialog(this, "SAP User account updated");
                    Fetch();
                    user_edit.setEnabled(false);
                    remove();
                } else {
                    JOptionPane.showMessageDialog(this, "SAP User account not updated");
                }
            } else {
                String duplicate = "SELECT * FROM user WHERE user_code = '" + code + "'";
                pst = con.prepareStatement(duplicate);
                rs = pst.executeQuery(duplicate);
                if (rs.next()) {
                    JOptionPane.showMessageDialog(this, "user code has been already used");
                    user_code.requestFocus();
                } else {
                    pst = con.prepareStatement("UPDATE user SET department=?,activity=?,db=?,user_code=?,name=?,license=?,remarks=? where user_code=?");
                    pst.setString(1, dp);
                    pst.setString(2, act);
                    pst.setString(3, db1);
                    pst.setString(4, code);
                    pst.setString(5, nm);
                    pst.setString(6, lt);
                    pst.setString(7, r);
                    pst.setString(8, user_code1.toString());
                    int k = pst.executeUpdate();

                    if (k == 1) {
                        JOptionPane.showMessageDialog(this, "SAP User account updated");
                        user_edit.setEnabled(false);
                        Fetch();
                        remove();
                    } else {
                        JOptionPane.showMessageDialog(this, "SAP User account not updated");
                    }
                }
            }

        } catch (SQLException ex) {
            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_user_editActionPerformed

    private void viber_tableMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_viber_tableMouseClicked
        int row = viber_table.getSelectedRow();
        Object client = viber_table.getModel().getValueAt(row, 0).toString();
        Object number = viber_table.getModel().getValueAt(row, 1).toString();
        Object dept = viber_table.getModel().getValueAt(row, 2).toString();
        Object device = viber_table.getModel().getValueAt(row, 3).toString();
        client_name.setText("" + client);
        mobile_number.setText("" + number);
        department.setText("" + dept);
        device_type.setSelectedItem(device);
        Delete.setEnabled(true);
        Viber_edit.setEnabled(true);
    }//GEN-LAST:event_viber_tableMouseClicked

    private void client_nameKeyReleased(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_client_nameKeyReleased

    }//GEN-LAST:event_client_nameKeyReleased

    private void mobile_numberKeyReleased(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_mobile_numberKeyReleased

    }//GEN-LAST:event_mobile_numberKeyReleased

    private void departmentKeyReleased(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_departmentKeyReleased

    }//GEN-LAST:event_departmentKeyReleased

    private void client_nameKeyTyped(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_client_nameKeyTyped
        Viber_add.setEnabled(true);
    }//GEN-LAST:event_client_nameKeyTyped

    private void mobile_numberKeyTyped(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_mobile_numberKeyTyped
        Viber_add.setEnabled(true);
    }//GEN-LAST:event_mobile_numberKeyTyped

    private void departmentKeyTyped(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_departmentKeyTyped
        Viber_add.setEnabled(true);
    }//GEN-LAST:event_departmentKeyTyped

    private void device_typeItemStateChanged(java.awt.event.ItemEvent evt) {//GEN-FIRST:event_device_typeItemStateChanged
        // TODO add your handling code here:
    }//GEN-LAST:event_device_typeItemStateChanged

    private void device_typeMouseReleased(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_device_typeMouseReleased
        // TODO add your handling code here:
    }//GEN-LAST:event_device_typeMouseReleased

    private void device_typeActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_device_typeActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_device_typeActionPerformed

    private void SaveDB21ActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_SaveDB21ActionPerformed
        String v = user_search.getText();

        if (v.isEmpty()) {
            JOptionPane.showMessageDialog(null, "You need to input first");
            DefaultTableModel df = (DefaultTableModel) search_table.getModel();
            df.setRowCount(0);
            isEmpty();
            user_search.requestFocus();
        } else {
            search_user();
            user_search.setText("");
        }
    }//GEN-LAST:event_SaveDB21ActionPerformed

    private void search_tableMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_search_tableMouseClicked
        int row = search_table.getSelectedRow();
        Object dep = search_table.getModel().getValueAt(row, 0).toString();
        Object act = search_table.getModel().getValueAt(row, 1).toString();
        Object db = search_table.getModel().getValueAt(row, 2).toString();
        Object code = search_table.getModel().getValueAt(row, 3).toString();
        Object name = search_table.getModel().getValueAt(row, 4).toString();
        Object license = search_table.getModel().getValueAt(row, 5).toString();
        Object rem = search_table.getModel().getValueAt(row, 6).toString();
        user_department.setText("" + dep);
        user_activity.setText("" + act);
        user_database.setText("" + db);
        user_code.setText("" + code);
        user_name.setText("" + name);
        LType.setSelectedItem(license);
        user_remark.setText("" + rem);
        user_remove.setEnabled(true);
        user_edit.setEnabled(true);
    }//GEN-LAST:event_search_tableMouseClicked

    private void jLabel96MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel96MouseClicked
    jTabbedPane1.setSelectedIndex(5);
          Fetch_Email();
        resetColor(user_tab);
        resetColor(search_tab);
        resetColor(home_tab);
        resetColor(inventory1);
        resetColor(com_tab);
        resetColor(lap_tab);
        resetColor(inv_search_tab);
        resetColor(viber_tab);
        setColor(email_tab);
        ViberTable(email_table);
       
        jLabel62.setForeground(new Color(228, 57, 39));
        jLabel48.setForeground(new Color(228, 57, 39));
        jLabel66.setForeground(new Color(228, 57, 39));
        jLabel55.setVisible(false);
        jLabel63.setVisible(false);
        jLabel73.setVisible(false);
        inv = "false";
    }//GEN-LAST:event_jLabel96MouseClicked

    private void jLabel97MouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_jLabel97MouseClicked
      jTabbedPane1.setSelectedIndex(5);
          Fetch_Email();
        resetColor(user_tab);
        resetColor(search_tab);
        resetColor(home_tab);
        resetColor(inventory1);
        resetColor(com_tab);
        resetColor(lap_tab);
        resetColor(inv_search_tab);
        resetColor(viber_tab);
        setColor(email_tab);
        ViberTable(email_table);
       
        jLabel62.setForeground(new Color(228, 57, 39));
        jLabel48.setForeground(new Color(228, 57, 39));
        jLabel66.setForeground(new Color(228, 57, 39));
        jLabel55.setVisible(false);
        jLabel63.setVisible(false);
        jLabel73.setVisible(false);
        inv = "false";
    }//GEN-LAST:event_jLabel97MouseClicked

    private void email_tabMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_email_tabMouseClicked
     jTabbedPane1.setSelectedIndex(5);

        resetColor(user_tab);
        resetColor(search_tab);
        resetColor(home_tab);
        resetColor(inventory1);
        resetColor(com_tab);
        resetColor(lap_tab);
        resetColor(inv_search_tab);
        resetColor(viber_tab);
        setColor(email_tab);
        ViberTable(email_table);
       Fetch_Email();
        jLabel62.setForeground(new Color(228, 57, 39));
        jLabel48.setForeground(new Color(228, 57, 39));
        jLabel66.setForeground(new Color(228, 57, 39));
        jLabel55.setVisible(false);
        jLabel63.setVisible(false);
        jLabel73.setVisible(false);
        inv = "false";
    }//GEN-LAST:event_email_tabMouseClicked

    private void email_tabMouseEntered(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_email_tabMouseEntered
        // TODO add your handling code here:
    }//GEN-LAST:event_email_tabMouseEntered

    private void email_tabMouseExited(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_email_tabMouseExited
        // TODO add your handling code here:
    }//GEN-LAST:event_email_tabMouseExited

    private void email_tabMousePressed(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_email_tabMousePressed
        // TODO add your handling code here:
    }//GEN-LAST:event_email_tabMousePressed

    private void email_nameKeyReleased(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_email_nameKeyReleased
        // TODO add your handling code here:
    }//GEN-LAST:event_email_nameKeyReleased

    private void email_nameKeyTyped(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_email_nameKeyTyped
        email_add.setEnabled(true);
    }//GEN-LAST:event_email_nameKeyTyped

    private void email_positionKeyReleased(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_email_positionKeyReleased
        // TODO add your handling code here:
    }//GEN-LAST:event_email_positionKeyReleased

    private void email_positionKeyTyped(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_email_positionKeyTyped
      email_add.setEnabled(true);
    }//GEN-LAST:event_email_positionKeyTyped

    private void email_emailKeyReleased(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_email_emailKeyReleased
        // TODO add your handling code here:
    }//GEN-LAST:event_email_emailKeyReleased

    private void email_emailKeyTyped(java.awt.event.KeyEvent evt) {//GEN-FIRST:event_email_emailKeyTyped
      email_add.setEnabled(true);
    }//GEN-LAST:event_email_emailKeyTyped

    private void email_departmentItemStateChanged(java.awt.event.ItemEvent evt) {//GEN-FIRST:event_email_departmentItemStateChanged
        // TODO add your handling code here:
    }//GEN-LAST:event_email_departmentItemStateChanged

    private void email_departmentMouseReleased(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_email_departmentMouseReleased
        // TODO add your handling code here:
    }//GEN-LAST:event_email_departmentMouseReleased

    private void email_departmentActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_email_departmentActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_email_departmentActionPerformed

    private void email_addActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_email_addActionPerformed
       int saved = 0;
        try {
            String name = email_name.getText();
            String position = email_position.getText();
            String department = email_department.getSelectedItem().toString();
            String email = email_email.getText();
            
            if (name.equals("") || (position.equals("")) || (email.equals(""))) {
                JOptionPane.showMessageDialog(null, "Please input all the necessary information");
            } else {
                String dup = "Select * from email_list where email='" + email + "'";
                pst = con.prepareStatement(dup);
                rs = pst.executeQuery(dup);

                if (rs.next()) {

                    JOptionPane.showMessageDialog(null, "Email has been already used.");
                    email_email.requestFocus();

                } else {
                    String dup1 = "Select * from email_list where name='" + name + "'";
                    pst = con.prepareStatement(dup1);
                    rs = pst.executeQuery(dup1);
                    if (rs.next()) {
                        JOptionPane.showMessageDialog(null, "Name has been already used.");
                        email_name.requestFocus();
                    } else {
                        pst = con.prepareStatement("Insert into email_list( name,position,department,email) values (?,?,?,?)");
                        pst.setString(1, name);
                        pst.setString(2, position);
                        pst.setString(3, department);
                        pst.setString(4, email);

                        saved = pst.executeUpdate();
                    }

                }
                if (saved == 1) {
                    JOptionPane.showMessageDialog(null, "SAVED!");
                    Fetch_Email();
                    email_name.setText("");
                    email_position.setText("");
                    email_email.setText("");
                    
                    email_department.setSelectedIndex(0);

                }
            }

        } catch (SQLException ex) {
            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);

        }
    }//GEN-LAST:event_email_addActionPerformed

    private void email_importActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_email_importActionPerformed
       int RowCount = email_table.getRowCount();
        if (RowCount != 0) {
            DefaultTableModel df = (DefaultTableModel) email_table.getModel();
            df.setRowCount(0);
            RowCount = 0;
            if (RowCount == 0) {
                Import_Email();
            }
        } else {
            Import_Email();
        }
    }//GEN-LAST:event_email_importActionPerformed

    private void email_exportActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_email_exportActionPerformed
        DefaultTableModel model = (DefaultTableModel) email_table.getModel();
        int rowCount = model.getRowCount();
        if (rowCount == 0) {
            JOptionPane.showMessageDialog(null, "You need to import files");
        } else {
            try {
                String defaultCurrentDirectoryPath = "C:\\Users\\Arnie D\\Downloads\\Arnie\\OJT\\IT DEPT";
                JFileChooser jFileChooser = new JFileChooser(defaultCurrentDirectoryPath);
                jFileChooser.showSaveDialog(this);
                File saveFile = jFileChooser.getSelectedFile();
                if (saveFile != null) {
                    saveFile = new File(saveFile.toString() + ".xlsx");
                    if (saveFile.exists()) {
                        int response = JOptionPane.showConfirmDialog(null, //
                                "Do you want to replace the existing file?", //
                                "Confirm", JOptionPane.YES_NO_OPTION, //
                                JOptionPane.QUESTION_MESSAGE);
                        if (response != JOptionPane.YES_OPTION) {

                            jFileChooser.showSaveDialog(this);
                             saveFile = jFileChooser.getSelectedFile();
                            saveFile = new File(saveFile.toString()+ ".xlsx");
                            Workbook wb = new XSSFWorkbook();
                            Sheet sheet = wb.createSheet("EMAIL LIST ACCOUNTS");
                            Row rowCol = sheet.createRow(0);
                            for (int c = 0; c < email_table.getColumnCount(); c++) {
                                Cell cell = rowCol.createCell(c);
                                cell.setCellValue(email_table.getColumnName(c));
                            }
                            for (int r = 0; r < email_table.getRowCount(); r++) {
                                Row row = sheet.createRow(r);
                                for (int rc = 0; rc < email_table.getColumnCount(); rc++) {
                                    Cell cell = row.createCell(rc);
                                    if (email_table.getValueAt(r, rc) != null) {
                                        cell.setCellValue(email_table.getValueAt(r, rc).toString());
                                    }
                                }
                            }
                            FileOutputStream out = new FileOutputStream(new File(saveFile.toString()));
                            wb.write(out);
                            wb.close();
                            out.close();
                            openFile(saveFile.toString());

                        }
                    } else if(!saveFile.exists()) {
                        Workbook wb = new XSSFWorkbook();
                        Sheet sheet = wb.createSheet("EMAIL LIST ACCOUNTS");
                        Row rowCol = sheet.createRow(0);
                        for (int c = 0; c < email_table.getColumnCount(); c++) {
                            Cell cell = rowCol.createCell(c);
                            cell.setCellValue(email_table.getColumnName(c));
                        }
                        for (int r = 0; r < email_table.getRowCount(); r++) {
                            Row row = sheet.createRow(r);
                            for (int rc = 0; rc < email_table.getColumnCount(); rc++) {
                                Cell cell = row.createCell(rc);
                                if (email_table.getValueAt(r, rc) != null) {
                                    cell.setCellValue(email_table.getValueAt(r, rc).toString());
                                }
                            }
                        }
                        FileOutputStream out = new FileOutputStream(new File(saveFile.toString()));
                        wb.write(out);
                        wb.close();
                        out.close();
                        openFile(saveFile.toString());
                    }
                } else {
                    JOptionPane.showMessageDialog(null, "Error!");
                }

            } catch (FileNotFoundException e) {
                System.out.println(e);
            } catch (IOException io) {
                System.out.println(io);
            }  
        }
    }//GEN-LAST:event_email_exportActionPerformed

    private void email_deleteActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_email_deleteActionPerformed
     try {
            DefaultTableModel model = (DefaultTableModel) email_table.getModel();
            if (email_table.getSelectedRowCount() == 0) {
                JOptionPane.showMessageDialog(this, "Account not selected");
            } else {
                int row = email_table.getSelectedRow();
                Object name = email_table.getModel().getValueAt(row, 0);
                Object position = email_table.getModel().getValueAt(row, 1);
                Object department = email_table.getModel().getValueAt(row, 2);
                Object email = email_table.getModel().getValueAt(row, 3);
                pst = con.prepareStatement("DELETE FROM email_list WHERE name=? and position=? and department=? and email=?");
                pst.setString(1, name.toString());
                pst.setString(2, position.toString());
                pst.setString(3, department.toString());
                pst.setString(4, email.toString());
                int k = pst.executeUpdate();
                if (k == 1) {
                    JOptionPane.showMessageDialog(this, "User deleted");
                    Fetch_Email();
                } else {
                    JOptionPane.showMessageDialog(this, "User not deleted");
                }
            }
        } catch (SQLException ex) {
            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_email_deleteActionPerformed

    private void email_editActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_email_editActionPerformed
        try {
            String name = email_name.getText();
            String pos = email_position.getText();
            String dept = email_department.getSelectedItem().toString(); 
            String ema = email_email.getText();
            int row = email_table.getSelectedRow();
            Object em = email_table.getModel().getValueAt(row, 3);

            if (ema.equals(em)) {
                pst = con.prepareStatement("UPDATE email_list SET name =?, position=?, department=?, email=? WHERE email=? ");
                pst.setString(1, name);
                pst.setString(2, pos);
                pst.setString(3, dept);
                pst.setString(4, ema);
                pst.setString(5, em.toString());
                int k = pst.executeUpdate();

                if (k == 1) {
                    JOptionPane.showMessageDialog(this, "Account updated");
                    Fetch_Email();
                } else {
                    JOptionPane.showMessageDialog(this, "Account not updated");
                }
            } else {
                String duplicate = "SELECT * FROM email_list WHERE email = '" + ema + "'";
                pst = con.prepareStatement(duplicate);
                rs = pst.executeQuery(duplicate);
                if (rs.next()) {
                    JOptionPane.showMessageDialog(this, "Email Acount already used");
                    email_email.requestFocus();
                } else {
                    pst = con.prepareStatement("UPDATE email_list SET name =?, position=?, department=?, email=? WHERE email=? ");
                    pst.setString(1, name);
                    pst.setString(2, pos);
                    pst.setString(3, dept);
                    pst.setString(4, ema);
                    pst.setString(5, em.toString());
                    int k = pst.executeUpdate();

                    if (k == 1) {
                        JOptionPane.showMessageDialog(this, "Account updated");
                        Fetch_Email();
                    } else {
                        JOptionPane.showMessageDialog(this, "Account noth updated");
                    }
                }
            }

        } catch (SQLException ex) {
            Logger.getLogger(Employee.class.getName()).log(Level.SEVERE, null, ex);
        }
    }//GEN-LAST:event_email_editActionPerformed

    private void email_printActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_email_printActionPerformed
      DefaultTableModel model = (DefaultTableModel) email_table.getModel();
        int rowCount = model.getRowCount();
        if (rowCount == 0) {
            JOptionPane.showMessageDialog(null, "You neeed to import files to print!");
        } else {
            MessageFormat header = new MessageFormat("Email List");
            MessageFormat footer = new MessageFormat("");
            try {
                
                email_table.print(JTable.PrintMode.FIT_WIDTH, header, footer);
                 
            } catch (PrinterException e) {
                JOptionPane.showMessageDialog(null, "Unable to Print!");
            }
        }
    }//GEN-LAST:event_email_printActionPerformed

    private void email_searchbuttonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_email_searchbuttonActionPerformed
       String v = email_search.getText();

        if (v.isEmpty()) {
            JOptionPane.showMessageDialog(null, "You need to input first");
            DefaultTableModel df = (DefaultTableModel) email_table.getModel();
            df.setRowCount(0);
            Fetch_Email();
            email_search.requestFocus();
        } else {
            search_email();
            search.setText("");
        }
    }//GEN-LAST:event_email_searchbuttonActionPerformed

    private void email_tableMouseClicked(java.awt.event.MouseEvent evt) {//GEN-FIRST:event_email_tableMouseClicked
       int row = email_table.getSelectedRow();
        Object name = email_table.getModel().getValueAt(row, 0).toString();
        Object position = email_table.getModel().getValueAt(row, 1).toString();
        Object dept = email_table.getModel().getValueAt(row, 2).toString();
        Object email = email_table.getModel().getValueAt(row, 3).toString();
        email_name.setText("" + name);
        email_position.setText("" + position);
        email_department.setSelectedItem(dept); 
        email_email.setText("" + email);
        email_delete.setEnabled(true);
        email_edit.setEnabled(true);
    }//GEN-LAST:event_email_tableMouseClicked

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
            java.util.logging.Logger.getLogger(Employee.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(Employee.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(Employee.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Employee.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
               
                new Employee().setVisible(true);

            }
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JComboBox<String> Assets;
    private javax.swing.JPanel COMPUTER;
    private javax.swing.JButton Delete;
    private javax.swing.JButton Export_Viber;
    private javax.swing.JComboBox<String> LType;
    private javax.swing.JButton SaveDB;
    private javax.swing.JButton SaveDB10;
    private javax.swing.JButton SaveDB11;
    private javax.swing.JButton SaveDB12;
    private javax.swing.JButton SaveDB13;
    private javax.swing.JButton SaveDB14;
    private javax.swing.JButton SaveDB2;
    private javax.swing.JButton SaveDB20;
    private javax.swing.JButton SaveDB21;
    private javax.swing.JButton SaveDB6;
    private javax.swing.JButton SaveDB9;
    private javax.swing.JButton Viber_Import;
    private javax.swing.JButton Viber_add;
    private javax.swing.JButton Viber_edit;
    private javax.swing.JTextField acct;
    private javax.swing.JTextField brnd;
    private javax.swing.JComboBox<String> c_d;
    private javax.swing.JTextField c_dept;
    private javax.swing.JButton c_export;
    private javax.swing.JComboBox<String> c_fb;
    private javax.swing.JTextField c_gc;
    private javax.swing.JComboBox<String> c_h;
    private javax.swing.JTextField c_hdb;
    private javax.swing.JTextField c_hds;
    private javax.swing.JTextField c_hdsn;
    private javax.swing.JButton c_import;
    private javax.swing.JTextField c_ip;
    private javax.swing.JTextField c_mab;
    private javax.swing.JTextField c_mak;
    private javax.swing.JTextField c_mb;
    private javax.swing.JTextField c_mbm;
    private javax.swing.JTextField c_mmrysn;
    private javax.swing.JTextField c_mn;
    private javax.swing.JTextField c_ms;
    private javax.swing.JTextField c_msn;
    private javax.swing.JTextField c_n;
    private javax.swing.JComboBox<String> c_oal;
    private javax.swing.JTextField c_p;
    private javax.swing.JButton c_print;
    private javax.swing.JTextField c_ps;
    private javax.swing.JTextField c_psb;
    private javax.swing.JTextField c_pssn;
    private javax.swing.JButton c_save;
    private javax.swing.JTextField c_sn;
    private javax.swing.JComboBox<String> c_usb;
    private javax.swing.JComboBox<String> c_w;
    private com.toedter.calendar.JDateChooser c_wedp;
    private javax.swing.JComboBox<String> c_wla;
    private javax.swing.JComboBox<String> c_yt;
    private javax.swing.JTextField client_name;
    private javax.swing.JPanel com_tab;
    private javax.swing.JTable computer_table;
    private javax.swing.JTextField condi;
    private com.toedter.calendar.JDateChooser date;
    private javax.swing.JTextField department;
    private javax.swing.JTextField dept1;
    private javax.swing.JTextField desc;
    private javax.swing.JComboBox<String> device_type;
    private javax.swing.JButton email_add;
    private javax.swing.JButton email_delete;
    private javax.swing.JComboBox<String> email_department;
    private javax.swing.JButton email_edit;
    private javax.swing.JTextField email_email;
    private javax.swing.JButton email_export;
    private javax.swing.JButton email_import;
    private javax.swing.JTextField email_name;
    private javax.swing.JTextField email_position;
    private javax.swing.JButton email_print;
    private javax.swing.JTextField email_search;
    private javax.swing.JButton email_searchbutton;
    private javax.swing.JPanel email_tab;
    private javax.swing.JTable email_table;
    private javax.swing.JButton export;
    private javax.swing.JPanel home_p1;
    private javax.swing.JPanel home_tab;
    private javax.swing.JTextField id;
    private javax.swing.JPanel inv_search_tab;
    private javax.swing.JTable inv_search_table;
    private javax.swing.JPanel inventory1;
    private javax.swing.JButton jButton1;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel10;
    private javax.swing.JLabel jLabel100;
    private javax.swing.JLabel jLabel101;
    private javax.swing.JLabel jLabel102;
    private javax.swing.JLabel jLabel103;
    private javax.swing.JLabel jLabel104;
    private javax.swing.JLabel jLabel105;
    private javax.swing.JLabel jLabel106;
    private javax.swing.JLabel jLabel107;
    private javax.swing.JLabel jLabel108;
    private javax.swing.JLabel jLabel109;
    private javax.swing.JLabel jLabel11;
    private javax.swing.JLabel jLabel110;
    private javax.swing.JLabel jLabel111;
    private javax.swing.JLabel jLabel112;
    private javax.swing.JLabel jLabel113;
    private javax.swing.JLabel jLabel114;
    private javax.swing.JLabel jLabel115;
    private javax.swing.JLabel jLabel116;
    private javax.swing.JLabel jLabel117;
    private javax.swing.JLabel jLabel118;
    private javax.swing.JLabel jLabel119;
    private javax.swing.JLabel jLabel12;
    private javax.swing.JLabel jLabel120;
    private javax.swing.JLabel jLabel121;
    private javax.swing.JLabel jLabel13;
    private javax.swing.JLabel jLabel14;
    private javax.swing.JLabel jLabel15;
    private javax.swing.JLabel jLabel16;
    private javax.swing.JLabel jLabel17;
    private javax.swing.JLabel jLabel18;
    private javax.swing.JLabel jLabel19;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel20;
    private javax.swing.JLabel jLabel21;
    private javax.swing.JLabel jLabel22;
    private javax.swing.JLabel jLabel23;
    private javax.swing.JLabel jLabel24;
    private javax.swing.JLabel jLabel25;
    private javax.swing.JLabel jLabel26;
    private javax.swing.JLabel jLabel27;
    private javax.swing.JLabel jLabel28;
    private javax.swing.JLabel jLabel29;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel30;
    private javax.swing.JLabel jLabel31;
    private javax.swing.JLabel jLabel32;
    private javax.swing.JLabel jLabel33;
    private javax.swing.JLabel jLabel34;
    private javax.swing.JLabel jLabel35;
    private javax.swing.JLabel jLabel36;
    private javax.swing.JLabel jLabel37;
    private javax.swing.JLabel jLabel38;
    private javax.swing.JLabel jLabel39;
    private javax.swing.JLabel jLabel40;
    private javax.swing.JLabel jLabel41;
    private javax.swing.JLabel jLabel42;
    private javax.swing.JLabel jLabel43;
    private javax.swing.JLabel jLabel44;
    private javax.swing.JLabel jLabel45;
    private javax.swing.JLabel jLabel46;
    private javax.swing.JLabel jLabel47;
    private javax.swing.JLabel jLabel48;
    private javax.swing.JLabel jLabel49;
    private javax.swing.JLabel jLabel5;
    private javax.swing.JLabel jLabel50;
    private javax.swing.JLabel jLabel51;
    private javax.swing.JLabel jLabel52;
    private javax.swing.JLabel jLabel53;
    private javax.swing.JLabel jLabel54;
    private javax.swing.JLabel jLabel55;
    private javax.swing.JLabel jLabel56;
    private javax.swing.JLabel jLabel57;
    private javax.swing.JLabel jLabel58;
    private javax.swing.JLabel jLabel59;
    private javax.swing.JLabel jLabel6;
    private javax.swing.JLabel jLabel60;
    private javax.swing.JLabel jLabel61;
    private javax.swing.JLabel jLabel62;
    private javax.swing.JLabel jLabel63;
    private javax.swing.JLabel jLabel64;
    private javax.swing.JLabel jLabel65;
    private javax.swing.JLabel jLabel66;
    private javax.swing.JLabel jLabel67;
    private javax.swing.JLabel jLabel68;
    private javax.swing.JLabel jLabel69;
    private javax.swing.JLabel jLabel7;
    private javax.swing.JLabel jLabel70;
    private javax.swing.JLabel jLabel71;
    private javax.swing.JLabel jLabel72;
    private javax.swing.JLabel jLabel73;
    private javax.swing.JLabel jLabel74;
    private javax.swing.JLabel jLabel75;
    private javax.swing.JLabel jLabel76;
    private javax.swing.JLabel jLabel77;
    private javax.swing.JLabel jLabel78;
    private javax.swing.JLabel jLabel79;
    private javax.swing.JLabel jLabel8;
    private javax.swing.JLabel jLabel80;
    private javax.swing.JLabel jLabel81;
    private javax.swing.JLabel jLabel82;
    private javax.swing.JLabel jLabel83;
    private javax.swing.JLabel jLabel84;
    private javax.swing.JLabel jLabel85;
    private javax.swing.JLabel jLabel86;
    private javax.swing.JLabel jLabel87;
    private javax.swing.JLabel jLabel88;
    private javax.swing.JLabel jLabel89;
    private javax.swing.JLabel jLabel9;
    private javax.swing.JLabel jLabel90;
    private javax.swing.JLabel jLabel91;
    private javax.swing.JLabel jLabel92;
    private javax.swing.JLabel jLabel93;
    private javax.swing.JLabel jLabel94;
    private javax.swing.JLabel jLabel95;
    private javax.swing.JLabel jLabel96;
    private javax.swing.JLabel jLabel97;
    private javax.swing.JLabel jLabel98;
    private javax.swing.JLabel jLabel99;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel10;
    private javax.swing.JPanel jPanel11;
    private javax.swing.JPanel jPanel12;
    private javax.swing.JPanel jPanel13;
    private javax.swing.JPanel jPanel14;
    private javax.swing.JPanel jPanel15;
    private javax.swing.JPanel jPanel16;
    private javax.swing.JPanel jPanel17;
    private javax.swing.JPanel jPanel18;
    private javax.swing.JPanel jPanel19;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JPanel jPanel20;
    private javax.swing.JPanel jPanel21;
    private javax.swing.JPanel jPanel22;
    private javax.swing.JPanel jPanel23;
    private javax.swing.JPanel jPanel24;
    private javax.swing.JPanel jPanel25;
    private javax.swing.JPanel jPanel26;
    private javax.swing.JPanel jPanel27;
    private javax.swing.JPanel jPanel28;
    private javax.swing.JPanel jPanel29;
    private javax.swing.JPanel jPanel3;
    private javax.swing.JPanel jPanel30;
    private javax.swing.JPanel jPanel31;
    private javax.swing.JPanel jPanel32;
    private javax.swing.JPanel jPanel33;
    private javax.swing.JPanel jPanel34;
    private javax.swing.JPanel jPanel35;
    private javax.swing.JPanel jPanel36;
    private javax.swing.JPanel jPanel37;
    private javax.swing.JPanel jPanel38;
    private javax.swing.JPanel jPanel39;
    private javax.swing.JPanel jPanel4;
    private javax.swing.JPanel jPanel40;
    private javax.swing.JPanel jPanel41;
    private javax.swing.JPanel jPanel5;
    private javax.swing.JPanel jPanel6;
    private javax.swing.JPanel jPanel7;
    private javax.swing.JPanel jPanel8;
    private javax.swing.JPanel jPanel9;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JScrollPane jScrollPane3;
    private javax.swing.JScrollPane jScrollPane5;
    private javax.swing.JScrollPane jScrollPane6;
    private javax.swing.JScrollPane jScrollPane7;
    private javax.swing.JScrollPane jScrollPane8;
    private javax.swing.JTabbedPane jTabbedPane1;
    public javax.swing.JTabbedPane jTabbedPane2;
    private javax.swing.JButton l_export;
    private javax.swing.JButton l_import;
    private javax.swing.JButton l_print;
    private javax.swing.JButton l_print1;
    private javax.swing.JButton l_print2;
    private javax.swing.JButton l_save;
    private javax.swing.JPanel lap_tab;
    private javax.swing.JTable laptop_table;
    private javax.swing.JTextField mdl;
    private javax.swing.JTextField mobile_number;
    private javax.swing.JTextField reco;
    private javax.swing.JTextField search;
    private javax.swing.JPanel search_tab;
    private javax.swing.JTable search_table;
    private javax.swing.JTextField srl;
    private javax.swing.JTextField status;
    private javax.swing.JTextField user_activity;
    private javax.swing.JTextField user_code;
    private javax.swing.JTextField user_database;
    private javax.swing.JTextField user_department;
    private javax.swing.JButton user_edit;
    private javax.swing.JTextField user_name;
    private javax.swing.JTextField user_remark;
    private javax.swing.JButton user_remove;
    private javax.swing.JTextField user_search;
    private javax.swing.JPanel user_tab;
    private javax.swing.JTable user_table;
    private javax.swing.JTextField viber_search;
    private javax.swing.JPanel viber_tab;
    private javax.swing.JTable viber_table;
    // End of variables declaration//GEN-END:variables
}
