
import com.formdev.flatlaf.themes.FlatMacLightLaf;

import javax.swing.*;
import javax.swing.border.Border;
import javax.swing.border.EmptyBorder;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import java.awt.BorderLayout;
import java.awt.CardLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Cursor;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.Frame;
import java.awt.Graphics;
import java.awt.Graphics2D;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.GridLayout;
import java.awt.Insets;
import java.awt.RenderingHints;
import java.awt.event.ActionEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AttendanceSystemUI {

    // --- Main Entry Point ---
    public static void main(String[] args) {
        // Use a modern macOS-inspired light theme
        FlatMacLightLaf.setup();

        // Global UI tweaks for a more modern look
        UIManager.put("Button.arc", 999); // Fully rounded buttons
        UIManager.put("Component.arc", 12);
        UIManager.put("ProgressBar.arc", 12);
        UIManager.put("TextComponent.arc", 8);
        UIManager.put("Table.showVerticalLines", false);
        UIManager.put("Table.intercellSpacing", new Dimension(0, 1));
        UIManager.put("TableHeader.height", 40);
        UIManager.put("TableHeader.background", new Color(245, 248, 251));
        UIManager.put("TableHeader.separatorColor", new Color(230, 230, 230));

        ExcelDataManager.setupDatabase();

        SwingUtilities.invokeLater(() -> {
            LoginFrame loginFrame = new LoginFrame();
            loginFrame.setVisible(true);
        });
    }

    // --- Modernized Color and Font Constants ---
    static class AppStyles {

        public static final Color BACKGROUND_COLOR = new Color(245, 248, 251);
        public static final Color SIDENAV_COLOR = Color.WHITE;
        public static final Color PRIMARY_COLOR = new Color(76, 110, 245);
        public static final Color PRIMARY_TEXT_COLOR = new Color(33, 37, 41);
        public static final Color SECONDARY_TEXT_COLOR = new Color(134, 142, 150);
        public static final Color GREEN = new Color(39, 174, 96);
        public static final Color RED = new Color(231, 76, 60);
        public static final Color BORDER_COLOR = new Color(222, 226, 230);
        public static final Color TABLE_ALT_ROW_COLOR = new Color(250, 251, 253);

        public static final Font FONT_BOLD = new Font("Inter", Font.BOLD, 14);
        public static final Font FONT_NORMAL = new Font("Inter", Font.PLAIN, 14);
        public static final Font FONT_HEADER = new Font("Inter", Font.BOLD, 28);
        public static final Font FONT_SMALL = new Font("Inter", Font.PLAIN, 12);
    }

    // --- Data Models (Unchanged) ---
    static class User {

        String id, password, name, role;

        User(String id, String password, String name, String role) {
            this.id = id;
            this.password = password;
            this.name = name;
            this.role = role;
        }
    }

    static class AttendanceRecord {

        String studentId, date, status;

        AttendanceRecord(String studentId, String date, String status) {
            this.studentId = studentId;
            this.date = date;
            this.status = status;
        }
    }

    // --- Main Application Frame ---
    static class MainFrame extends JFrame {

        private final CardLayout cardLayout = new CardLayout();
        private final JPanel contentPanel = new JPanel(cardLayout);
        private User currentUser;
        private final SideNavPanel sideNavPanel;

        MainFrame(User user) {
            this.currentUser = user;
            setTitle("Attendance System");
            setSize(1280, 800);
            setMinimumSize(new Dimension(1100, 700));
            setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            setLocationRelativeTo(null);
            setLayout(new BorderLayout());
            getContentPane().setBackground(AppStyles.BACKGROUND_COLOR);

            sideNavPanel = new SideNavPanel(this);
            add(sideNavPanel, BorderLayout.WEST);

            setupContentPanels();
            add(contentPanel, BorderLayout.CENTER);

            // Set initial view based on role
            switch (currentUser.role) {
                case "Admin":
                    sideNavPanel.setActive("Students");
                    showPanel("Admin");
                    break;
                case "Staff":
                    sideNavPanel.setActive("Attendance");
                    showPanel("Staff");
                    break;
                case "Student":
                    sideNavPanel.setActive("Dashboard");
                    showPanel("Student");
                    break;
            }
        }

        private void setupContentPanels() {
            contentPanel.add(new StaffAttendancePanel(this), "Staff");
            contentPanel.add(new AdminPanel(this), "Admin");
            contentPanel.add(new StudentDashboardPanel(this), "Student");
        }

        public void showPanel(String panelName) {
            cardLayout.show(contentPanel, panelName);
            Component[] components = contentPanel.getComponents();
            for (Component component : components) {
                if (component.getClass().getSimpleName().equals(panelName + "Panel") && component.isVisible()) {
                    if (component instanceof Refreshable) {
                        ((Refreshable) component).refreshData();
                    }
                    break;
                }
            }
        }

        public User getCurrentUser() {
            return currentUser;
        }
    }

    // --- Redesigned Login Frame ---
    static class LoginFrame extends JFrame {

        LoginFrame() {
            setTitle("Login");
            setSize(800, 600);
            setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            setLocationRelativeTo(null);
            setLayout(new GridLayout(1, 2));
            setResizable(false);

            // Left Panel (Branding)
            JPanel brandPanel = new JPanel(new GridBagLayout());
            brandPanel.setBackground(AppStyles.PRIMARY_COLOR);

            JLabel brandTitle = new JLabel("<html>College<br>Attendance<br>System</html>");
            brandTitle.setFont(new Font("Inter", Font.BOLD, 48));
            brandTitle.setForeground(Color.WHITE);
            brandPanel.add(brandTitle);
            add(brandPanel);

            // Right Panel (Form)
            JPanel formPanel = new JPanel(new GridBagLayout());
            formPanel.setBackground(Color.WHITE);
            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(10, 25, 10, 25);
            gbc.gridwidth = 1;
            gbc.fill = GridBagConstraints.HORIZONTAL;

            JLabel title = new JLabel("Welcome Back!");
            title.setFont(AppStyles.FONT_HEADER);
            gbc.gridx = 0;
            gbc.gridy = 0;
            gbc.weightx = 1.0;
            formPanel.add(title, gbc);

            JLabel subtitle = new JLabel("Please enter your details to sign in.");
            subtitle.setFont(AppStyles.FONT_NORMAL);
            subtitle.setForeground(AppStyles.SECONDARY_TEXT_COLOR);
            gbc.gridy = 1;
            gbc.insets = new Insets(0, 25, 20, 25);
            formPanel.add(subtitle, gbc);

            JTextField usernameField = new JTextField(20);
            JPasswordField passwordField = new JPasswordField(20);
            setupTextField(usernameField, "Full Name");
            setupTextField(passwordField, "Password");

            gbc.gridy = 2;
            gbc.insets = new Insets(10, 25, 10, 25);
            formPanel.add(usernameField, gbc);
            gbc.gridy = 3;
            formPanel.add(passwordField, gbc);

            JButton loginButton = new JButton("Login");
            loginButton.setBackground(AppStyles.PRIMARY_COLOR);
            loginButton.setForeground(Color.WHITE);
            loginButton.setFont(AppStyles.FONT_BOLD);
            loginButton.setPreferredSize(new Dimension(100, 45));
            gbc.gridy = 4;
            gbc.insets = new Insets(20, 25, 10, 25);
            formPanel.add(loginButton, gbc);

            loginButton.addActionListener(e -> performLogin(usernameField.getText(), new String(passwordField.getPassword())));
            add(formPanel);
        }

        private void setupTextField(JTextField field, String placeholder) {
            field.putClientProperty("JTextField.placeholderText", placeholder);
            field.setPreferredSize(new Dimension(200, 45));
        }

        private void performLogin(String username, String password) {
            if (username.isEmpty() || password.isEmpty()) {
                JOptionPane.showMessageDialog(this, "Fields cannot be empty.", "Error", JOptionPane.ERROR_MESSAGE);
                return;
            }
            // MODIFIED: Authenticate with username (which is the user's name)
            User user = ExcelDataManager.authenticateUser(username, password);
            if (user != null) {
                dispose();
                new MainFrame(user).setVisible(true);
            } else {
                JOptionPane.showMessageDialog(this, "Invalid credentials.", "Login Failed", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    // --- Improved Side Navigation Panel ---
    static class SideNavPanel extends JPanel {

        private final MainFrame mainFrame;
        private final JPanel buttonsPanel = new JPanel();
        private JButton activeButton;

        SideNavPanel(MainFrame frame) {
            this.mainFrame = frame;
            setLayout(new BorderLayout(0, 30));
            setPreferredSize(new Dimension(240, 0));
            setBackground(AppStyles.SIDENAV_COLOR);
            setBorder(BorderFactory.createMatteBorder(0, 0, 0, 1, AppStyles.BORDER_COLOR));

            JLabel titleLabel = new JLabel("ATTENDANCE SYS", SwingConstants.CENTER);
            titleLabel.setFont(new Font("Inter", Font.BOLD, 20));
            titleLabel.setForeground(AppStyles.PRIMARY_TEXT_COLOR);
            titleLabel.setBorder(new EmptyBorder(25, 10, 25, 10));
            add(titleLabel, BorderLayout.NORTH);

            buttonsPanel.setLayout(new GridLayout(0, 1, 0, 15));
            buttonsPanel.setOpaque(false);
            buttonsPanel.setBorder(new EmptyBorder(15, 15, 15, 15));

            String role = mainFrame.getCurrentUser().role;
            switch (role) {
                case "Admin":
                    addNavButton("👤", "Students", "Admin");
                    addNavButton("📄", "Reports", "Admin"); // Simplified to show same panel
                    break;
                case "Staff":
                    addNavButton("✅", "Attendance", "Staff");
                    break;
                case "Student":
                    addNavButton("📊", "Dashboard", "Student");
                    break;
            }

            add(buttonsPanel, BorderLayout.CENTER);

            JButton logoutButton = new JButton("Logout");
            logoutButton.setFont(AppStyles.FONT_BOLD);
            logoutButton.putClientProperty("JButton.buttonType", "roundRect");
            logoutButton.addActionListener(e -> {
                mainFrame.dispose();
                new LoginFrame().setVisible(true);
            });
            JPanel logoutWrapper = new JPanel(new FlowLayout(FlowLayout.CENTER));
            logoutWrapper.setOpaque(false);
            logoutWrapper.add(logoutButton);
            logoutWrapper.setBorder(new EmptyBorder(0, 0, 20, 0));
            add(logoutWrapper, BorderLayout.SOUTH);
        }

        private void addNavButton(String icon, String text, String panelName) {
            JButton button = new JButton(icon + "    " + text);
            button.setFont(AppStyles.FONT_BOLD);
            button.setFocusPainted(false);
            button.setHorizontalAlignment(SwingConstants.LEFT);
            button.setBorder(new EmptyBorder(12, 25, 12, 25));
            button.setCursor(new Cursor(Cursor.HAND_CURSOR));
            setInactiveStyle(button);

            button.addActionListener(e -> {
                setActive(text);
                mainFrame.showPanel(panelName);
            });
            buttonsPanel.add(button);
        }

        public void setActive(String text) {
            if (activeButton != null) {
                setInactiveStyle(activeButton);
            }
            for (Component comp : buttonsPanel.getComponents()) {
                if (comp instanceof JButton && ((JButton) comp).getText().contains(text)) {
                    activeButton = (JButton) comp;
                    setActiveStyle(activeButton);
                    break;
                }
            }
        }

        private void setActiveStyle(JButton button) {
            button.setBackground(AppStyles.PRIMARY_COLOR);
            button.setForeground(Color.WHITE);
        }

        private void setInactiveStyle(JButton button) {
            button.setBackground(AppStyles.SIDENAV_COLOR);
            button.setForeground(AppStyles.SECONDARY_TEXT_COLOR);
        }
    }

    // --- Reusable UI Components with Modern Styling ---
    static class CustomComponents {

        public static JPanel createCardPanel() {
            JPanel card = new JPanel();
            card.setBackground(Color.WHITE);
            card.setBorder(BorderFactory.createCompoundBorder(
                    BorderFactory.createLineBorder(AppStyles.BORDER_COLOR, 1, true),
                    new EmptyBorder(15, 15, 15, 15)
            ));
            return card;
        }

        public static JPanel createStatCard(String title, String value, String icon, Color color) {
            JPanel card = createCardPanel();
            card.setLayout(new BorderLayout(15, 0));

            JLabel iconLabel = new JLabel(icon);
            iconLabel.setFont(new Font("Segoe UI Symbol", Font.PLAIN, 32));
            iconLabel.setForeground(color);
            iconLabel.setVerticalAlignment(SwingConstants.CENTER);
            card.add(iconLabel, BorderLayout.WEST);

            JPanel textPanel = new JPanel();
            textPanel.setOpaque(false);
            textPanel.setLayout(new BoxLayout(textPanel, BoxLayout.Y_AXIS));

            JLabel valueLabel = new JLabel(value);
            valueLabel.setFont(new Font("Inter", Font.BOLD, 26));
            valueLabel.setForeground(AppStyles.PRIMARY_TEXT_COLOR);

            JLabel titleLabel = new JLabel(title);
            titleLabel.setFont(AppStyles.FONT_SMALL);
            titleLabel.setForeground(AppStyles.SECONDARY_TEXT_COLOR);

            textPanel.add(valueLabel);
            textPanel.add(Box.createRigidArea(new Dimension(0, 5)));
            textPanel.add(titleLabel);
            card.add(textPanel, BorderLayout.CENTER);
            return card;
        }

        public static class StatusCellRenderer extends DefaultTableCellRenderer {

            @Override
            public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
                JPanel panel = new JPanel(new FlowLayout(FlowLayout.CENTER, 0, 0));
                JLabel label = new JLabel(value.toString());
                label.setFont(new Font("Inter", Font.BOLD, 12));
                label.setBorder(new EmptyBorder(6, 18, 6, 18));

                if ("Present".equals(value)) {
                    label.setBackground(new Color(230, 245, 237));
                    label.setForeground(AppStyles.GREEN.darker());
                } else if ("Absent".equals(value)) {
                    label.setBackground(new Color(253, 235, 232));
                    label.setForeground(AppStyles.RED.darker());
                }

                label.setOpaque(true);
                panel.setOpaque(true);
                panel.setBackground(isSelected ? table.getSelectionBackground() : table.getBackground());

                return panel;
            }
        }

        public static JTable createModernTable(DefaultTableModel model) {
            JTable table = new JTable(model);
            table.setRowHeight(45);
            table.setFont(AppStyles.FONT_NORMAL);
            table.setForeground(AppStyles.PRIMARY_TEXT_COLOR);
            table.setSelectionBackground(AppStyles.PRIMARY_COLOR.brighter());
            table.setSelectionForeground(Color.WHITE);
            table.getTableHeader().setFont(AppStyles.FONT_BOLD);
            table.getTableHeader().setForeground(AppStyles.SECONDARY_TEXT_COLOR);
            table.putClientProperty("JTable.alternateRowColor", AppStyles.TABLE_ALT_ROW_COLOR);
            return table;
        }
    }

    interface Refreshable {

        void refreshData();
    }

    // --- Modernized Staff View Panel ---
    static class StaffAttendancePanel extends JPanel implements Refreshable {

        private final DefaultTableModel tableModel;
        private final JPanel statsPanel = new JPanel(new GridLayout(1, 4, 20, 20));

        StaffAttendancePanel(MainFrame frame) {
            setLayout(new BorderLayout(20, 20));
            setBorder(new EmptyBorder(25, 25, 25, 25));
            setBackground(AppStyles.BACKGROUND_COLOR);

            // Top Header
            JPanel headerPanel = new JPanel(new BorderLayout());
            headerPanel.setOpaque(false);
            JLabel title = new JLabel("Attendance");
            title.setFont(AppStyles.FONT_HEADER);
            headerPanel.add(title, BorderLayout.WEST);
            add(headerPanel, BorderLayout.NORTH);

            // Stats Panel
            statsPanel.setOpaque(false);
            add(statsPanel, BorderLayout.CENTER);

            // Main Content Panel for table
            JPanel mainContent = CustomComponents.createCardPanel();
            mainContent.setLayout(new BorderLayout(15, 15));
            mainContent.setBorder(new EmptyBorder(15, 15, 15, 15));

            // Table Toolbar
            JPanel toolbar = new JPanel(new FlowLayout(FlowLayout.LEFT));
            toolbar.setOpaque(false);
            toolbar.add(new JLabel("Today's Roster: " + new SimpleDateFormat("EEE, d MMM yyyy").format(new Date())));
            toolbar.getComponent(0).setFont(AppStyles.FONT_BOLD);
            mainContent.add(toolbar, BorderLayout.NORTH);

            // Table
            String[] columns = {"Roll No.", "Name", "Status"};
            tableModel = new DefaultTableModel(columns, 0) {
                @Override
                public boolean isCellEditable(int row, int column) {
                    return false;
                }
            };
            JTable table = CustomComponents.createModernTable(tableModel);
            table.getColumnModel().getColumn(2).setCellRenderer(new CustomComponents.StatusCellRenderer());
            table.getColumnModel().getColumn(0).setPreferredWidth(50);
            table.getColumnModel().getColumn(1).setPreferredWidth(200);

            table.addMouseListener(new MouseAdapter() {
                public void mouseClicked(MouseEvent e) {
                    int row = table.rowAtPoint(e.getPoint());
                    int col = table.columnAtPoint(e.getPoint());
                    if (row >= 0 && col == 2) {
                        String currentStatus = (String) tableModel.getValueAt(row, 2);
                        String newStatus = "Present".equals(currentStatus) ? "Absent" : "Present";
                        tableModel.setValueAt(newStatus, row, 2);
                        updateStats();
                    }
                }
            });

            JScrollPane scrollPane = new JScrollPane(table);
            scrollPane.setBorder(BorderFactory.createLineBorder(AppStyles.BORDER_COLOR));
            mainContent.add(scrollPane, BorderLayout.CENTER);

            // Bottom Actions
            JPanel actionsPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 10, 0));
            actionsPanel.setOpaque(false);
            JButton markAllPresent = new JButton("Mark All Present");
            JButton markAllAbsent = new JButton("Mark All Absent");
            JButton saveButton = new JButton("Save Attendance");
            saveButton.setBackground(AppStyles.PRIMARY_COLOR);
            saveButton.setForeground(Color.WHITE);
            saveButton.setFont(AppStyles.FONT_BOLD);
            actionsPanel.add(markAllPresent);
            actionsPanel.add(markAllAbsent);
            actionsPanel.add(saveButton);

            mainContent.add(actionsPanel, BorderLayout.SOUTH);

            markAllPresent.addActionListener(e -> setAllStatus("Present"));
            markAllAbsent.addActionListener(e -> setAllStatus("Absent"));
            saveButton.addActionListener(this::submitAttendance);

            // This structure is a bit different to put stats above table
            JPanel centerPanel = new JPanel(new BorderLayout(20, 20));
            centerPanel.setOpaque(false);
            centerPanel.add(statsPanel, BorderLayout.NORTH);
            centerPanel.add(mainContent, BorderLayout.CENTER);
            add(centerPanel, BorderLayout.CENTER);

            refreshData();
        }

        @Override
        public void refreshData() {
            tableModel.setRowCount(0);
            List<User> students = ExcelDataManager.getUsersByRole("Student");
            for (User student : students) {
                String todayStatus = ExcelDataManager.getStudentStatusForToday(student.id);
                tableModel.addRow(new Object[]{student.id, student.name, todayStatus});
            }
            updateStats();
        }

        private void updateStats() {
            int present = 0, absent = 0;
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                if ("Present".equals(tableModel.getValueAt(i, 2))) {
                    present++; 
                }else {
                    absent++;
                }
            }
            int total = tableModel.getRowCount();
            double percentage = (total == 0) ? 0 : ((double) present / total) * 100;

            statsPanel.removeAll();
            statsPanel.add(CustomComponents.createStatCard("Present Today", String.valueOf(present), "✅", AppStyles.GREEN));
            statsPanel.add(CustomComponents.createStatCard("Absent Today", String.valueOf(absent), "❌", AppStyles.RED));
            statsPanel.add(CustomComponents.createStatCard("Attendance %", String.format("%.0f%%", percentage), "📈", AppStyles.PRIMARY_COLOR));
            statsPanel.add(CustomComponents.createStatCard("Total Students", String.valueOf(total), "🎓", Color.ORANGE));
            statsPanel.revalidate();
            statsPanel.repaint();
        }

        private void setAllStatus(String status) {
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                tableModel.setValueAt(status, i, 2);
            }
            updateStats();
        }

        private void submitAttendance(ActionEvent e) {
            List<AttendanceRecord> records = new ArrayList<>();
            String date = new SimpleDateFormat("yyyy-MM-dd").format(new Date());
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                records.add(new AttendanceRecord((String) tableModel.getValueAt(i, 0), date, (String) tableModel.getValueAt(i, 2)));
            }
            if (ExcelDataManager.hasAttendanceBeenMarkedToday()) {
                int choice = JOptionPane.showConfirmDialog(this,
                        "Overwrite today's attendance record?", "Confirm", JOptionPane.YES_NO_OPTION);
                if (choice == JOptionPane.NO_OPTION) {
                    return;
                }
            }
            ExcelDataManager.markAttendance(records);
            JOptionPane.showMessageDialog(this, "Attendance saved successfully!", "Success", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    // --- Modernized Admin Panel ---
    static class AdminPanel extends JPanel implements Refreshable {

        private final JTable studentTable, staffTable, reportTable;
        private final DefaultTableModel studentModel, staffModel, reportModel;

        AdminPanel(MainFrame frame) {
            setLayout(new BorderLayout(20, 20));
            setBorder(new EmptyBorder(25, 25, 25, 25));
            setBackground(AppStyles.BACKGROUND_COLOR);

            JLabel title = new JLabel("Admin Dashboard");
            title.setFont(AppStyles.FONT_HEADER);
            add(title, BorderLayout.NORTH);

            JTabbedPane tabbedPane = new JTabbedPane();
            tabbedPane.setFont(AppStyles.FONT_BOLD);

            studentModel = new DefaultTableModel(new String[]{"ID", "Name", "Role"}, 0) {
                @Override
                public boolean isCellEditable(int row, int column) {
                    return false;
                }
            };
            studentTable = CustomComponents.createModernTable(studentModel);
            tabbedPane.addTab("Manage Students", createManagementPanel(studentTable, "Student", frame, this::refreshData));

            staffModel = new DefaultTableModel(new String[]{"ID", "Name", "Role"}, 0) {
                @Override
                public boolean isCellEditable(int row, int column) {
                    return false;
                }
            };
            staffTable = CustomComponents.createModernTable(staffModel);
            tabbedPane.addTab("Manage Staff", createManagementPanel(staffTable, "Staff", frame, this::refreshData));

            reportModel = new DefaultTableModel(new String[]{"Student ID", "Name", "Date", "Status"}, 0);
            reportTable = CustomComponents.createModernTable(reportModel);
            reportTable.getColumnModel().getColumn(3).setCellRenderer(new CustomComponents.StatusCellRenderer());
            tabbedPane.addTab("Attendance Report", new JScrollPane(reportTable));

            add(tabbedPane, BorderLayout.CENTER);
            refreshData();
        }

        private JPanel createManagementPanel(JTable table, String role, Frame owner, Runnable refreshCallback) {
            JPanel panel = new JPanel(new BorderLayout(10, 10));
            panel.add(new JScrollPane(table), BorderLayout.CENTER);

            JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
            JButton addButton = new JButton("➕ Add New " + role);
            JButton removeButton = new JButton("❌ Remove Selected");
            addButton.setFont(AppStyles.FONT_BOLD);
            removeButton.setFont(AppStyles.FONT_BOLD);
            addButton.setBackground(AppStyles.GREEN);
            addButton.setForeground(Color.WHITE);

            addButton.addActionListener(e -> new AddUserDialog(owner, role, refreshCallback).setVisible(true));
            removeButton.addActionListener(e -> {
                int selectedRow = table.getSelectedRow();
                if (selectedRow != -1) {
                    String id = (String) table.getModel().getValueAt(table.convertRowIndexToModel(selectedRow), 0);
                    int confirm = JOptionPane.showConfirmDialog(this, "Are you sure you want to remove user " + id + "?", "Confirm Deletion", JOptionPane.YES_NO_OPTION);
                    if (confirm == JOptionPane.YES_OPTION) {
                        ExcelDataManager.removeUser(id);
                        refreshCallback.run();
                    }
                } else {
                    JOptionPane.showMessageDialog(this, "Please select a user to remove.", "Warning", JOptionPane.WARNING_MESSAGE);
                }
            });

            buttonPanel.add(addButton);
            buttonPanel.add(removeButton);
            panel.add(buttonPanel, BorderLayout.SOUTH);
            return panel;
        }

        @Override
        public void refreshData() {
            studentModel.setRowCount(0);
            ExcelDataManager.getUsersByRole("Student").forEach(u -> studentModel.addRow(new Object[]{u.id, u.name, u.role}));
            staffModel.setRowCount(0);
            ExcelDataManager.getUsersByRole("Staff").forEach(u -> staffModel.addRow(new Object[]{u.id, u.name, u.role}));
            reportModel.setRowCount(0);
            ExcelDataManager.getAllAttendance().forEach(r -> {
                User student = ExcelDataManager.getUserById(r.studentId);
                reportModel.addRow(new Object[]{r.studentId, student != null ? student.name : "N/A", r.date, r.status});
            });
        }
    }

    // --- Modernized Student Dashboard ---
    static class StudentDashboardPanel extends JPanel implements Refreshable {

        private final DefaultTableModel tableModel;
        private final JPanel statsPanel = new JPanel(new GridLayout(1, 3, 20, 20));

        StudentDashboardPanel(MainFrame frame) {
            setLayout(new BorderLayout(20, 20));
            setBorder(new EmptyBorder(25, 25, 25, 25));
            setBackground(AppStyles.BACKGROUND_COLOR);

            JLabel title = new JLabel("Welcome, " + frame.getCurrentUser().name);
            title.setFont(AppStyles.FONT_HEADER);
            add(title, BorderLayout.NORTH);

            statsPanel.setOpaque(false);

            tableModel = new DefaultTableModel(new String[]{"Date", "Status"}, 0) {
                @Override
                public boolean isCellEditable(int row, int column) {
                    return false;
                }
            };
            JTable table = CustomComponents.createModernTable(tableModel);
            table.getColumnModel().getColumn(1).setCellRenderer(new CustomComponents.StatusCellRenderer());

            JPanel tableCard = CustomComponents.createCardPanel();
            tableCard.setLayout(new BorderLayout());
            JLabel historyLabel = new JLabel("Attendance History");
            historyLabel.setFont(AppStyles.FONT_BOLD);
            historyLabel.setBorder(new EmptyBorder(5, 5, 10, 5));
            tableCard.add(historyLabel, BorderLayout.NORTH);
            tableCard.add(new JScrollPane(table), BorderLayout.CENTER);

            JPanel centerPanel = new JPanel(new BorderLayout(20, 20));
            centerPanel.setOpaque(false);
            centerPanel.add(statsPanel, BorderLayout.NORTH);
            centerPanel.add(tableCard, BorderLayout.CENTER);
            add(centerPanel, BorderLayout.CENTER);

            refreshData(frame.getCurrentUser());
        }

        public void refreshData(User currentUser) {
            tableModel.setRowCount(0);
            List<AttendanceRecord> records = ExcelDataManager.getAttendanceForStudent(currentUser.id);
            int present = 0, absent = 0;
            for (AttendanceRecord r : records) {
                tableModel.addRow(new Object[]{r.date, r.status});
                if ("Present".equals(r.status)) {
                    present++;
                } else {
                    absent++;
                }
            }
            int total = records.size();
            double percentage = (total == 0) ? 0 : ((double) present / total) * 100;

            statsPanel.removeAll();
            statsPanel.add(CustomComponents.createStatCard("Total Present", String.valueOf(present), "✅", AppStyles.GREEN));
            statsPanel.add(CustomComponents.createStatCard("Total Absent", String.valueOf(absent), "❌", AppStyles.RED));
            statsPanel.add(CustomComponents.createStatCard("Overall Percentage", String.format("%.0f%%", percentage), "📈", AppStyles.PRIMARY_COLOR));
            statsPanel.revalidate();
            statsPanel.repaint();
        }

        @Override
        public void refreshData() {
            /* Not used directly */ }
    }

    // --- Modernized Add User Dialog ---
    static class AddUserDialog extends JDialog {

        AddUserDialog(Frame owner, String role, Runnable refreshCallback) {
            super(owner, "Add New " + role, true);
            setSize(400, 250); // Adjusted size
            setLocationRelativeTo(owner);
            setLayout(new GridBagLayout());

            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(8, 8, 8, 8);
            gbc.fill = GridBagConstraints.HORIZONTAL;

            // MODIFIED: Removed the ID field
            JTextField nameField = new JTextField();
            JPasswordField passField = new JPasswordField();

            gbc.gridx = 0;
            gbc.gridy = 0;
            add(new JLabel("Full Name:"), gbc);
            gbc.gridx = 1;
            gbc.gridy = 0;
            gbc.weightx = 1.0;
            add(nameField, gbc);
            gbc.gridx = 0;
            gbc.gridy = 1;
            add(new JLabel("Password:"), gbc);
            gbc.gridx = 1;
            gbc.gridy = 1;
            add(passField, gbc);

            JButton addButton = new JButton("Add User");
            addButton.setBackground(AppStyles.PRIMARY_COLOR);
            addButton.setForeground(Color.WHITE);
            addButton.setFont(AppStyles.FONT_BOLD);
            gbc.gridx = 1;
            gbc.gridy = 2; // Adjusted position
            gbc.anchor = GridBagConstraints.EAST;
            gbc.fill = GridBagConstraints.NONE;
            add(addButton, gbc);

            addButton.addActionListener(e -> {
                // MODIFIED: Call new addUser method
                ExcelDataManager.addUser(nameField.getText(), new String(passField.getPassword()), role);
                refreshCallback.run();
                dispose();
            });
        }
    }

    // --- Excel Data Manager (with updated logic) ---
    static class ExcelDataManager {

        private static final String FILE_NAME = "college_data.xlsx";
        private static final String USERS_SHEET = "Users";
        private static final String ATTENDANCE_SHEET = "Attendance";

        public static void setupDatabase() {
            if (!new File(FILE_NAME).exists()) {
                try (Workbook workbook = new XSSFWorkbook()) {
                    Sheet usersSheet = workbook.createSheet(USERS_SHEET);
                    Row header = usersSheet.createRow(0);
                    header.createCell(0).setCellValue("ID");
                    header.createCell(1).setCellValue("Password");
                    header.createCell(2).setCellValue("Name");
                    header.createCell(3).setCellValue("Role");
                    Row adminRow = usersSheet.createRow(1);
                    adminRow.createCell(0).setCellValue("admin");
                    adminRow.createCell(1).setCellValue("admin123");
                    adminRow.createCell(2).setCellValue("Administrator"); // Name for admin login
                    adminRow.createCell(3).setCellValue("Admin");

                    Sheet attendanceSheet = workbook.createSheet(ATTENDANCE_SHEET);
                    Row attHeader = attendanceSheet.createRow(0);
                    attHeader.createCell(0).setCellValue("StudentID");
                    attHeader.createCell(1).setCellValue("Date");
                    attHeader.createCell(2).setCellValue("Status");

                    try (FileOutputStream fos = new FileOutputStream(FILE_NAME)) {
                        workbook.write(fos);
                    }
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        // MODIFIED: Authenticates using Name (column 2) instead of ID
        public static User authenticateUser(String username, String password) {
            try (FileInputStream fis = new FileInputStream(FILE_NAME); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheet(USERS_SHEET);
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) {
                        continue;
                    }
                    // Check name in column 2 (index 2)
                    if (row.getCell(2) != null && Objects.equals(row.getCell(2).getStringCellValue(), username) && Objects.equals(row.getCell(1).getStringCellValue(), password)) {
                        return new User(row.getCell(0).getStringCellValue(), password, username, row.getCell(3).getStringCellValue());
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
            return null;
        }

        // MODIFIED: Generates unique ID automatically
        public static void addUser(String name, String password, String role) {
            String idPrefix = role.equalsIgnoreCase("Student") ? "STU" : "STAFF";

            List<User> usersOfRole = getUsersByRole(role);
            int lastIdNum = usersOfRole.stream()
                    .map(u -> u.id.substring(idPrefix.length()))
                    .mapToInt(Integer::parseInt)
                    .max()
                    .orElse(0);

            String newId = String.format("%s%03d", idPrefix, lastIdNum + 1);
            User newUser = new User(newId, password, name, role);

            try (FileInputStream fis = new FileInputStream(FILE_NAME); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheet(USERS_SHEET);
                Row newRow = sheet.createRow(sheet.getLastRowNum() + 1);
                newRow.createCell(0).setCellValue(newUser.id);
                newRow.createCell(1).setCellValue(newUser.password);
                newRow.createCell(2).setCellValue(newUser.name);
                newRow.createCell(3).setCellValue(newUser.role);
                try (FileOutputStream fos = new FileOutputStream(FILE_NAME)) {
                    workbook.write(fos);
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        public static void removeUser(String id) {
            try (FileInputStream fis = new FileInputStream(FILE_NAME); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheet(USERS_SHEET);
                int rowToRemove = -1;
                for (Row row : sheet) {
                    if (row.getRowNum() == 0 || row.getCell(0) == null) {
                        continue;
                    }
                    if (row.getCell(0).getStringCellValue().equalsIgnoreCase(id)) {
                        rowToRemove = row.getRowNum();
                        break;
                    }
                }
                if (rowToRemove != -1) {
                    sheet.removeRow(sheet.getRow(rowToRemove));
                    if (rowToRemove <= sheet.getLastRowNum()) {
                        sheet.shiftRows(rowToRemove + 1, sheet.getLastRowNum(), -1);
                    }
                }
                try (FileOutputStream fos = new FileOutputStream(FILE_NAME)) {
                    workbook.write(fos);
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        public static List<User> getUsersByRole(String role) {
            List<User> users = new ArrayList<>();
            try (FileInputStream fis = new FileInputStream(FILE_NAME); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheet(USERS_SHEET);
                for (Row row : sheet) {
                    if (row.getRowNum() == 0 || row.getCell(3) == null) {
                        continue;
                    }
                    if (row.getCell(3).getStringCellValue().equalsIgnoreCase(role)) {
                        users.add(new User(row.getCell(0).getStringCellValue(), "", row.getCell(2).getStringCellValue(), role));
                    }
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
            return users;
        }

        public static User getUserById(String userId) {
            try (FileInputStream fis = new FileInputStream(FILE_NAME); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheet(USERS_SHEET);
                for (Row row : sheet) {
                    if (row.getRowNum() == 0 || row.getCell(0) == null) {
                        continue;
                    }
                    if (row.getCell(0).getStringCellValue().equalsIgnoreCase(userId)) {
                        return new User(row.getCell(0).getStringCellValue(), "", row.getCell(2).getStringCellValue(), row.getCell(3).getStringCellValue());
                    }
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
            return null;
        }

        public static List<AttendanceRecord> getAllAttendance() {
            List<AttendanceRecord> records = new ArrayList<>();
            try (FileInputStream fis = new FileInputStream(FILE_NAME); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheet(ATTENDANCE_SHEET);
                for (Row row : sheet) {
                    if (row.getRowNum() == 0 || row.getCell(0) == null) {
                        continue;
                    }
                    records.add(new AttendanceRecord(row.getCell(0).getStringCellValue(), row.getCell(1).getStringCellValue(), row.getCell(2).getStringCellValue()));
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
            return records;
        }

        public static List<AttendanceRecord> getAttendanceForStudent(String studentId) {
            List<AttendanceRecord> records = new ArrayList<>();
            for (AttendanceRecord r : getAllAttendance()) {
                if (r.studentId.equalsIgnoreCase(studentId)) {
                    records.add(r);
                }
            }
            return records;
        }

        public static boolean hasAttendanceBeenMarkedToday() {
            String today = new SimpleDateFormat("yyyy-MM-dd").format(new Date());
            return getAllAttendance().stream().anyMatch(r -> r.date.equals(today));
        }

        public static String getStudentStatusForToday(String studentId) {
            String today = new SimpleDateFormat("yyyy-MM-dd").format(new Date());
            return getAllAttendance().stream()
                    .filter(r -> r.date.equals(today) && r.studentId.equals(studentId))
                    .map(r -> r.status)
                    .findFirst().orElse("Present"); // Default to present
        }

        public static void markAttendance(List<AttendanceRecord> records) {
            try (FileInputStream fis = new FileInputStream(FILE_NAME); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheet(ATTENDANCE_SHEET);
                String today = new SimpleDateFormat("yyyy-MM-dd").format(new Date());

                List<Integer> rowsToRemove = new ArrayList<>();
                for (Row row : sheet) {
                    if (row.getRowNum() > 0 && row.getCell(1) != null && row.getCell(1).getStringCellValue().equals(today)) {
                        rowsToRemove.add(row.getRowNum());
                    }
                }
                for (int i = rowsToRemove.size() - 1; i >= 0; i--) {
                    int rowIndex = rowsToRemove.get(i);
                    Row row = sheet.getRow(rowIndex);
                    if (row != null) {
                        sheet.removeRow(row);
                    }
                }
                if (!rowsToRemove.isEmpty()) {
                    int firstRow = rowsToRemove.get(0);
                    if (firstRow <= sheet.getLastRowNum()) {
                        sheet.shiftRows(firstRow, sheet.getLastRowNum(), -rowsToRemove.size());
                    }
                }

                for (AttendanceRecord record : records) {
                    Row newRow = sheet.createRow(sheet.getLastRowNum() + 1);
                    newRow.createCell(0).setCellValue(record.studentId);
                    newRow.createCell(1).setCellValue(record.date);
                    newRow.createCell(2).setCellValue(record.status);
                }

                try (FileOutputStream fos = new FileOutputStream(FILE_NAME)) {
                    workbook.write(fos);
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
