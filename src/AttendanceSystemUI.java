
import com.formdev.flatlaf.themes.FlatMacLightLaf;

import javax.swing.*;
import javax.swing.border.Border;
import javax.swing.border.EmptyBorder;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableModel;
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
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AttendanceSystemUI {

    // --- Centralized List of Subjects ---
    public static final String[] SUBJECT_LIST = {"General", "Math", "Physics", "Chemistry", "History", "English", "Biology"};
    private static final DateTimeFormatter GLOBAL_DATE_FORMATTER = DateTimeFormatter.ISO_LOCAL_DATE;

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

    // --- Data Models ---
    static class User {

        String id, password, name, role, subject;

        User(String id, String password, String name, String role) {
            this(id, password, name, role, ""); // Default subject to empty string
        }

        User(String id, String password, String name, String role, String subject) {
            this.id = id;
            this.password = password;
            this.name = name;
            this.role = role;
            this.subject = (subject == null) ? "" : subject;
        }
    }

    static class AttendanceRecord {

        String studentId, date, status, subject;

        AttendanceRecord(String studentId, String date, String status, String subject) {
            this.studentId = studentId;
            this.date = date;
            this.status = status;
            this.subject = subject;
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

            // NEW: Updated Login routing
            switch (currentUser.role) {
                case "Admin":
                case "Staff":
                    // Admin and Staff now land on the new Home Dashboard
                    sideNavPanel.setActive("Dashboard");
                    showPanel("HomeDashboard");
                    break;
                case "Student":
                    // Student still lands on their personal dashboard
                    sideNavPanel.setActive("Dashboard");
                    showPanel("StudentDashboard");
                    break;
            }
        }

        private void setupContentPanels() {
            contentPanel.add(new HomeDashboardPanel(this), "HomeDashboard"); // NEW: Admin/Staff Dashboard
            contentPanel.add(new StaffAttendancePanel(this), "Staff");
            contentPanel.add(new AdminPanel(this), "Admin");
            contentPanel.add(new StudentDashboardPanel(this), "StudentDashboard"); // Renamed key
            contentPanel.add(new ReportsPanel(this), "Reports");
        }

        public void showPanel(String panelName) {
            cardLayout.show(contentPanel, panelName);
            Component[] components = contentPanel.getComponents();
            for (Component component : components) {
                // Match class simple name OR the card layout key
                String compName = component.getClass().getSimpleName();
                if ((compName.equals(panelName) || compName.equals(panelName + "Panel"))
                        && component.isVisible()) {
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
            GridBagConstraints gbcBrand = new GridBagConstraints();
            gbcBrand.insets = new Insets(10, 10, 10, 10);

            JLabel brandIcon = new JLabel("🎓");
            brandIcon.setFont(new Font("Segoe UI Symbol", Font.PLAIN, 120));
            brandIcon.setForeground(Color.WHITE.darker());
            gbcBrand.gridy = 0;
            brandPanel.add(brandIcon, gbcBrand);

            JLabel brandTitle = new JLabel("<html><div style='text-align: center;'>College<br>Attendance<br>System</div></html>");
            brandTitle.setFont(new Font("Inter", Font.BOLD, 48));
            brandTitle.setForeground(Color.WHITE);
            gbcBrand.gridy = 1;
            brandPanel.add(brandTitle, gbcBrand);
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
            User user = ExcelDataManager.authenticateUser(username, password);
            if (user != null) {
                dispose();
                new MainFrame(user).setVisible(true);
            } else {
                JOptionPane.showMessageDialog(this, "Invalid credentials.", "Login Failed", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    // --- Improved Side Navigation Panel (Now with Dashboard) ---
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
                    addNavButton("🏠", "Dashboard", "HomeDashboard");
                    addNavButton("👤", "Students", "Admin"); // This panel is now just for management
                    addNavButton("📄", "Reports", "Reports");
                    break;
                case "Staff":
                    addNavButton("🏠", "Dashboard", "HomeDashboard");
                    addNavButton("✅", "Attendance", "Staff");
                    addNavButton("📄", "Reports", "Reports");
                    break;
                case "Student":
                    addNavButton("📊", "Dashboard", "StudentDashboard");
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
                Component c = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
                if (value == null) {
                    return c;
                }

                JPanel panel = new JPanel(new FlowLayout(FlowLayout.CENTER, 8, 0));
                JLabel label = new JLabel(value.toString());
                label.setFont(new Font("Inter", Font.BOLD, 12));
                label.setBorder(new EmptyBorder(6, 18, 6, 18));
                label.setOpaque(true);

                if ("Present".equals(value)) {
                    label.setBackground(new Color(230, 245, 237));
                    label.setForeground(AppStyles.GREEN.darker());
                } else if ("Absent".equals(value)) {
                    label.setBackground(new Color(253, 235, 232));
                    label.setForeground(AppStyles.RED.darker());
                } else {
                    label.setBackground(Color.WHITE);
                    label.setForeground(AppStyles.PRIMARY_TEXT_COLOR);
                    label.setBorder(null);
                }

                panel.add(label);
                panel.setOpaque(true);

                if (isSelected) {
                    panel.setBackground(table.getSelectionBackground());
                } else {
                    Color bg = (row % 2 == 0) ? Color.WHITE : AppStyles.TABLE_ALT_ROW_COLOR;
                    panel.setBackground(bg);
                    if (!"Present".equals(value) && !"Absent".equals(value)) {
                        label.setBackground(bg);
                    }
                }

                return panel;
            }
        }

        public static class PercentageCellRenderer extends DefaultTableCellRenderer {

            @Override
            public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
                Component c = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
                if (c instanceof JLabel) {
                    setHorizontalAlignment(SwingConstants.RIGHT);
                    if (value instanceof Number) {
                        ((JLabel) c).setText(String.format("%.0f%%", ((Number) value).doubleValue()));
                    }
                }
                if (isSelected) {
                    c.setBackground(table.getSelectionBackground());
                    c.setForeground(table.getSelectionForeground());
                } else {
                    c.setBackground((row % 2 == 0) ? Color.WHITE : AppStyles.TABLE_ALT_ROW_COLOR);
                    c.setForeground(table.getForeground());
                }
                return c;
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

            table.setDefaultRenderer(Object.class, new DefaultTableCellRenderer() {
                @Override
                public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
                    Component c = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
                    if (!isSelected) {
                        c.setBackground(row % 2 == 0 ? Color.WHITE : AppStyles.TABLE_ALT_ROW_COLOR);
                    }
                    return c;
                }
            });
            table.setOpaque(false);
            return table;
        }
    }

    interface Refreshable {

        void refreshData();
    }

    // --- NEW: Admin/Staff Home Dashboard Panel ---
    static class HomeDashboardPanel extends JPanel implements Refreshable {

        private final MainFrame mainFrame;
        private final User currentUser;
        // MODIFIED: Grid layout is now 1 row, 4 columns
        private final JPanel statsPanel = new JPanel(new GridLayout(1, 4, 20, 20));
        private final JLabel titleLabel;

        HomeDashboardPanel(MainFrame frame) {
            this.mainFrame = frame;
            this.currentUser = frame.getCurrentUser();

            setLayout(new BorderLayout(20, 20));
            setBorder(new EmptyBorder(25, 25, 25, 25));
            setBackground(AppStyles.BACKGROUND_COLOR);

            titleLabel = new JLabel("Welcome, " + currentUser.name);
            titleLabel.setFont(AppStyles.FONT_HEADER);
            add(titleLabel, BorderLayout.NORTH);

            statsPanel.setOpaque(false);

            JPanel centerMessage = new JPanel(new GridBagLayout());
            centerMessage.setOpaque(false);
            JLabel message = new JLabel("<html><div style='text-align: center;'>Select an option from the menu on the left to get started.</div></html>");
            message.setFont(AppStyles.FONT_NORMAL);
            message.setForeground(AppStyles.SECONDARY_TEXT_COLOR);
            centerMessage.add(message);

            JPanel centerPanel = new JPanel(new BorderLayout(20, 20));
            centerPanel.setOpaque(false);
            centerPanel.add(statsPanel, BorderLayout.NORTH);
            centerPanel.add(centerMessage, BorderLayout.CENTER);

            add(centerPanel, BorderLayout.CENTER);

            refreshData();
        }

        @Override
        public void refreshData() {
            statsPanel.removeAll();

            if ("Admin".equals(currentUser.role)) {
                List<User> students = ExcelDataManager.getUsersByRole("Student");
                List<User> staff = ExcelDataManager.getUsersByRole("Staff");
                Map<String, Double> todayStats = ExcelDataManager.getOverallAttendanceForToday();

                statsPanel.add(CustomComponents.createStatCard("Total Students", String.valueOf(students.size()), "🎓", AppStyles.PRIMARY_COLOR));
                statsPanel.add(CustomComponents.createStatCard("Total Staff", String.valueOf(staff.size()), "👤", Color.ORANGE));
                // NEW: Added Total Lectures card
                statsPanel.add(CustomComponents.createStatCard("Total Lectures Today", String.valueOf(todayStats.get("total").intValue()), "📚", AppStyles.GREEN));
                statsPanel.add(CustomComponents.createStatCard("Overall Absent Today", String.valueOf(todayStats.get("absent").intValue()), "❌", AppStyles.RED));

            } else if ("Staff".equals(currentUser.role)) {
                Map<String, Double> subjectStats = ExcelDataManager.getOverallAttendanceForToday(currentUser.subject);

                statsPanel.add(CustomComponents.createStatCard("My Assigned Subject", currentUser.subject, "📚", AppStyles.PRIMARY_COLOR));
                statsPanel.add(CustomComponents.createStatCard("Att. % (My Subject)", String.format("%.0f%%", subjectStats.get("percentage")), "📈", AppStyles.GREEN));
                // NEW: Added Total Lectures card
                statsPanel.add(CustomComponents.createStatCard("Total Lectures (My Subject)", String.valueOf(subjectStats.get("total").intValue()), "✍️", Color.ORANGE));
                statsPanel.add(CustomComponents.createStatCard("Absent (My Subject)", String.valueOf(subjectStats.get("absent").intValue()), "❌", AppStyles.RED));
            }

            statsPanel.revalidate();
            statsPanel.repaint();
        }
    }

    // --- Staff View Panel ---
    static class StaffAttendancePanel extends JPanel implements Refreshable {

        private final DefaultTableModel tableModel;
        private final JPanel statsPanel = new JPanel(new GridLayout(1, 4, 20, 20));
        private final JTextField dateField;
        private final String assignedSubject;
        private final JLabel subjectLabel;

        StaffAttendancePanel(MainFrame frame) {
            this.assignedSubject = frame.getCurrentUser().subject;

            setLayout(new BorderLayout(20, 20));
            setBorder(new EmptyBorder(25, 25, 25, 25));
            setBackground(AppStyles.BACKGROUND_COLOR);

            JPanel headerPanel = new JPanel(new BorderLayout());
            headerPanel.setOpaque(false);
            JLabel title = new JLabel("Attendance Manager");
            title.setFont(AppStyles.FONT_HEADER);
            headerPanel.add(title, BorderLayout.WEST);
            add(headerPanel, BorderLayout.NORTH);

            statsPanel.setOpaque(false);

            JPanel mainContent = CustomComponents.createCardPanel();
            mainContent.setLayout(new BorderLayout(15, 15));
            mainContent.setBorder(new EmptyBorder(15, 15, 15, 15));

            JPanel toolbar = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 0));
            toolbar.setOpaque(false);

            toolbar.add(new JLabel("Date:"));
            dateField = new JTextField(LocalDate.now().format(GLOBAL_DATE_FORMATTER), 12);
            dateField.putClientProperty("JTextField.placeholderText", "yyyy-MM-dd");
            toolbar.add(dateField);

            subjectLabel = new JLabel("Subject: " + this.assignedSubject);
            subjectLabel.setFont(AppStyles.FONT_BOLD);
            toolbar.add(subjectLabel);

            JButton loadButton = new JButton("Load Roster");
            loadButton.setFont(AppStyles.FONT_BOLD);
            loadButton.addActionListener(e -> loadRosterForAssignedSubject());
            toolbar.add(loadButton);

            mainContent.add(toolbar, BorderLayout.NORTH);

            String[] columns = {"Roll No.", "Name", "Status"};
            tableModel = new DefaultTableModel(columns, 0) {
                @Override
                public boolean isCellEditable(int r, int c) {
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
            scrollPane.getViewport().setBackground(Color.WHITE);
            scrollPane.setBorder(BorderFactory.createLineBorder(AppStyles.BORDER_COLOR));
            mainContent.add(scrollPane, BorderLayout.CENTER);

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

            JPanel centerPanel = new JPanel(new BorderLayout(20, 20));
            centerPanel.setOpaque(false);
            centerPanel.add(statsPanel, BorderLayout.NORTH);
            centerPanel.add(mainContent, BorderLayout.CENTER);
            add(centerPanel, BorderLayout.CENTER);

            refreshData();
        }

        private void loadRosterForAssignedSubject() {
            String dateStr;
            try {
                LocalDate.parse(dateField.getText(), GLOBAL_DATE_FORMATTER);
                dateStr = dateField.getText();
            } catch (DateTimeParseException e) {
                JOptionPane.showMessageDialog(this, "Invalid date format. Please use yyyy-MM-dd.", "Date Error", JOptionPane.ERROR_MESSAGE);
                return;
            }

            tableModel.setRowCount(0);
            List<User> students = ExcelDataManager.getUsersByRole("Student");
            for (User student : students) {
                String status = ExcelDataManager.getStudentStatusForDate(student.id, dateStr, this.assignedSubject);
                tableModel.addRow(new Object[]{student.id, student.name, status});
            }
            updateStats();
        }

        @Override
        public void refreshData() {
            dateField.setText(LocalDate.now().format(GLOBAL_DATE_FORMATTER));
            loadRosterForAssignedSubject();
        }

        private void updateStats() {
            int present = 0, absent = 0;
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                if ("Present".equals(tableModel.getValueAt(i, 2))) {
                    present++;
                } else {
                    absent++;
                }
            }
            int total = tableModel.getRowCount();
            double percentage = (total == 0) ? 100.0 : ((double) present / total) * 100;

            statsPanel.removeAll();
            statsPanel.add(CustomComponents.createStatCard("Present", String.valueOf(present), "✅", AppStyles.GREEN));
            statsPanel.add(CustomComponents.createStatCard("Absent", String.valueOf(absent), "❌", AppStyles.RED));
            statsPanel.add(CustomComponents.createStatCard("Total Roster", String.valueOf(total), "🎓", Color.ORANGE));
            statsPanel.add(CustomComponents.createStatCard("Attendance %", String.format("%.0f%%", percentage), "📈", AppStyles.PRIMARY_COLOR));
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
            String dateStr;
            try {
                LocalDate.parse(dateField.getText(), GLOBAL_DATE_FORMATTER);
                dateStr = dateField.getText();
            } catch (DateTimeParseException ex) {
                JOptionPane.showMessageDialog(this, "Invalid date format. Cannot save.", "Date Error", JOptionPane.ERROR_MESSAGE);
                return;
            }

            List<AttendanceRecord> records = new ArrayList<>();
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                records.add(new AttendanceRecord(
                        (String) tableModel.getValueAt(i, 0),
                        dateStr,
                        (String) tableModel.getValueAt(i, 2),
                        this.assignedSubject
                ));
            }

            if (ExcelDataManager.hasAttendanceBeenMarked(dateStr, this.assignedSubject)) {
                int choice = JOptionPane.showConfirmDialog(this,
                        "Overwrite attendance for '" + this.assignedSubject + "' on " + dateStr + "?",
                        "Confirm Overwrite", JOptionPane.YES_NO_OPTION);
                if (choice == JOptionPane.NO_OPTION) {
                    return;
                }
            }
            ExcelDataManager.markAttendance(records, dateStr, this.assignedSubject);
            JOptionPane.showMessageDialog(this, "Attendance for " + this.assignedSubject + " on " + dateStr + " saved successfully!", "Success", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    // --- Admin Panel (NOW WITH PASSWORD RESET) ---
    static class AdminPanel extends JPanel implements Refreshable {

        private final JTable studentTable, staffTable, reportTable;
        private final DefaultTableModel studentModel, staffModel, reportModel;

        AdminPanel(MainFrame frame) {
            setLayout(new BorderLayout(20, 20));
            setBorder(new EmptyBorder(25, 25, 25, 25));
            setBackground(AppStyles.BACKGROUND_COLOR);

            JLabel title = new JLabel("Admin Management");
            title.setFont(AppStyles.FONT_HEADER);
            add(title, BorderLayout.NORTH);

            JTabbedPane tabbedPane = new JTabbedPane();
            tabbedPane.setFont(AppStyles.FONT_BOLD);

            studentModel = new DefaultTableModel(new String[]{"ID", "Name", "Role"}, 0) {
                @Override
                public boolean isCellEditable(int r, int c) {
                    return false;
                }
            };
            studentTable = CustomComponents.createModernTable(studentModel);
            tabbedPane.addTab("Manage Students", createManagementPanel(studentTable, "Student", frame, this::refreshData));

            staffModel = new DefaultTableModel(new String[]{"ID", "Name", "Role", "Assigned Subject"}, 0) {
                @Override
                public boolean isCellEditable(int r, int c) {
                    return false;
                }
            };
            staffTable = CustomComponents.createModernTable(staffModel);
            tabbedPane.addTab("Manage Staff", createManagementPanel(staffTable, "Staff", frame, this::refreshData));

            reportModel = new DefaultTableModel(new String[]{"Student ID", "Name", "Date", "Subject", "Status"}, 0);
            reportTable = CustomComponents.createModernTable(reportModel);
            reportTable.getColumnModel().getColumn(4).setCellRenderer(new CustomComponents.StatusCellRenderer());
            tabbedPane.addTab("Raw Attendance Data", new JScrollPane(reportTable));

            add(tabbedPane, BorderLayout.CENTER);
            refreshData();
        }

        private JPanel createManagementPanel(JTable table, String role, Frame owner, Runnable refreshCallback) {
            JPanel panel = new JPanel(new BorderLayout(10, 10));
            panel.add(new JScrollPane(table), BorderLayout.CENTER);

            JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
            JButton addButton = new JButton("➕ Add New " + role);
            JButton removeButton = new JButton("❌ Remove Selected");
            JButton setPasswordButton = new JButton("🔑 Set Password"); // NEW BUTTON

            addButton.setFont(AppStyles.FONT_BOLD);
            removeButton.setFont(AppStyles.FONT_BOLD);
            setPasswordButton.setFont(AppStyles.FONT_BOLD);

            addButton.setBackground(AppStyles.GREEN);
            addButton.setForeground(Color.WHITE);
            setPasswordButton.setBackground(Color.ORANGE);

            addButton.addActionListener(e -> new AddUserDialog(owner, role, refreshCallback).setVisible(true));

            removeButton.addActionListener(e -> {
                int selectedRow = table.getSelectedRow();
                if (selectedRow != -1) {
                    String id = (String) table.getModel().getValueAt(table.convertRowIndexToModel(selectedRow), 0);
                    String name = (String) table.getModel().getValueAt(table.convertRowIndexToModel(selectedRow), 1);
                    int confirm = JOptionPane.showConfirmDialog(this, "Are you sure you want to remove user: " + name + " (ID: " + id + ")?", "Confirm Deletion", JOptionPane.YES_NO_OPTION, JOptionPane.WARNING_MESSAGE);
                    if (confirm == JOptionPane.YES_OPTION) {
                        ExcelDataManager.removeUser(id);
                        refreshCallback.run();
                    }
                } else {
                    JOptionPane.showMessageDialog(this, "Please select a user to remove.", "Warning", JOptionPane.WARNING_MESSAGE);
                }
            });

            setPasswordButton.addActionListener(e -> {
                int selectedRow = table.getSelectedRow();
                if (selectedRow != -1) {
                    String id = (String) table.getModel().getValueAt(table.convertRowIndexToModel(selectedRow), 0);
                    String name = (String) table.getModel().getValueAt(table.convertRowIndexToModel(selectedRow), 1);

                    String newPassword = JOptionPane.showInputDialog(this, "Enter new password for " + name + " (ID: " + id + "):", "Set New Password", JOptionPane.PLAIN_MESSAGE);

                    if (newPassword != null && !newPassword.trim().isEmpty()) {
                        ExcelDataManager.updatePassword(id, newPassword.trim());
                        JOptionPane.showMessageDialog(this, "Password for " + name + " updated successfully.", "Success", JOptionPane.INFORMATION_MESSAGE);
                    } else if (newPassword != null) {
                        JOptionPane.showMessageDialog(this, "Password cannot be empty.", "Error", JOptionPane.ERROR_MESSAGE);
                    }
                } else {
                    JOptionPane.showMessageDialog(this, "Please select a user to update.", "Warning", JOptionPane.WARNING_MESSAGE);
                }
            });

            buttonPanel.add(addButton);
            buttonPanel.add(setPasswordButton);
            buttonPanel.add(removeButton);
            panel.add(buttonPanel, BorderLayout.SOUTH);
            return panel;
        }

        @Override
        public void refreshData() {
            studentModel.setRowCount(0);
            ExcelDataManager.getUsersByRole("Student").forEach(u -> studentModel.addRow(new Object[]{u.id, u.name, u.role}));

            staffModel.setRowCount(0);
            ExcelDataManager.getUsersByRole("Staff").forEach(u -> staffModel.addRow(new Object[]{u.id, u.name, u.role, u.subject}));

            reportModel.setRowCount(0);
            ExcelDataManager.getAllAttendance().forEach(r -> {
                User student = ExcelDataManager.getUserById(r.studentId);
                reportModel.addRow(new Object[]{r.studentId, student != null ? student.name : "N/A", r.date, r.subject, r.status});
            });
        }
    }

    // --- STUDENT DASHBOARD (NOW WITH SUBJECT SUMMARY TAB) ---
    static class StudentDashboardPanel extends JPanel implements Refreshable {

        private final DefaultTableModel historyTableModel;
        private final DefaultTableModel subjectSummaryModel; // NEW: Table model for the summary
        // MODIFIED: Stat panel now 4 columns
        private final JPanel statsPanel = new JPanel(new GridLayout(1, 4, 20, 20));
        private final MainFrame mainFrame;

        StudentDashboardPanel(MainFrame frame) {
            this.mainFrame = frame;
            setLayout(new BorderLayout(20, 20));
            setBorder(new EmptyBorder(25, 25, 25, 25));
            setBackground(AppStyles.BACKGROUND_COLOR);

            JLabel title = new JLabel("Welcome, " + frame.getCurrentUser().name);
            title.setFont(AppStyles.FONT_HEADER);
            add(title, BorderLayout.NORTH);

            statsPanel.setOpaque(false);

            // --- NEW: Create Tabbed Pane ---
            JTabbedPane tabbedPane = new JTabbedPane();
            tabbedPane.setFont(AppStyles.FONT_BOLD);

            // --- Tab 1: Subject Summary (NEW) ---
            subjectSummaryModel = new DefaultTableModel(new String[]{"Subject", "Total Lectures", "Present", "Absent", "Percentage"}, 0) {
                @Override
                public boolean isCellEditable(int r, int c) {
                    return false;
                }
            };
            JTable subjectTable = CustomComponents.createModernTable(subjectSummaryModel);
            subjectTable.getColumnModel().getColumn(4).setCellRenderer(new CustomComponents.PercentageCellRenderer());
            tabbedPane.addTab("Subject Summary", new JScrollPane(subjectTable));

            // --- Tab 2: Recent History (Existing) ---
            historyTableModel = new DefaultTableModel(new String[]{"Date", "Subject", "Status"}, 0) {
                @Override
                public boolean isCellEditable(int r, int c) {
                    return false;
                }
            };
            JTable historyTable = CustomComponents.createModernTable(historyTableModel);
            historyTable.getColumnModel().getColumn(2).setCellRenderer(new CustomComponents.StatusCellRenderer());
            tabbedPane.addTab("Recent History", new JScrollPane(historyTable));

            // --- Layout Panel ---
            JPanel centerPanel = new JPanel(new BorderLayout(20, 20));
            centerPanel.setOpaque(false);
            centerPanel.add(statsPanel, BorderLayout.NORTH);
            centerPanel.add(tabbedPane, BorderLayout.CENTER); // Add tabbed pane instead of single card
            add(centerPanel, BorderLayout.CENTER);

            refreshData();
        }

        public void refreshData(User currentUser) {
            // Clear both tables
            historyTableModel.setRowCount(0);
            subjectSummaryModel.setRowCount(0);

            List<AttendanceRecord> records = ExcelDataManager.getAttendanceForStudent(currentUser.id);
            int overallPresent = 0;
            int overallAbsent = 0;

            // NEW: Map to build the subject summary
            // Use LinkedHashMap to maintain insertion order (looks cleaner)
            Map<String, int[]> subjectSummaryData = new LinkedHashMap<>();

            for (AttendanceRecord r : records) {
                // 1. Populate the history table (same as before)
                historyTableModel.addRow(new Object[]{r.date, r.subject, r.status});

                // 2. Update overall stats (same as before)
                if ("Present".equals(r.status)) {
                    overallPresent++;
                } else {
                    overallAbsent++;
                }

                // 3. NEW: Build the summary map
                // Get the count array for this subject, or create it if it's the first time
                int[] counts = subjectSummaryData.computeIfAbsent(r.subject, k -> new int[2]); // {present, absent}
                if ("Present".equals(r.status)) {
                    counts[0]++; // Increment present count
                } else {
                    counts[1]++; // Increment absent count
                }
            }

            int overallTotal = overallPresent + overallAbsent;
            double overallPercentage = (overallTotal == 0) ? 100.0 : ((double) overallPresent / overallTotal) * 100.0;

            // --- Update Top Stat Cards (Now with 4 stats) ---
            statsPanel.removeAll();
            statsPanel.add(CustomComponents.createStatCard("Total Lectures", String.valueOf(overallTotal), "📚", Color.ORANGE));
            statsPanel.add(CustomComponents.createStatCard("Total Present", String.valueOf(overallPresent), "✅", AppStyles.GREEN));
            statsPanel.add(CustomComponents.createStatCard("Total Absent", String.valueOf(overallAbsent), "❌", AppStyles.RED));
            statsPanel.add(CustomComponents.createStatCard("Overall Percentage", String.format("%.0f%%", overallPercentage), "📈", AppStyles.PRIMARY_COLOR));
            statsPanel.revalidate();
            statsPanel.repaint();

            // --- NEW: Populate the Subject Summary Table ---
            for (Map.Entry<String, int[]> entry : subjectSummaryData.entrySet()) {
                String subject = entry.getKey();
                int present = entry.getValue()[0];
                int absent = entry.getValue()[1];
                int total = present + absent;
                double percentage = (total == 0) ? 100.0 : ((double) present / total) * 100.0;

                subjectSummaryModel.addRow(new Object[]{subject, total, present, absent, percentage});
            }
        }

        @Override
        public void refreshData() {
            refreshData(mainFrame.getCurrentUser());
        }
    }

    // --- Advanced Reports Panel (BUG FIX Applied) ---
    static class ReportsPanel extends JPanel implements Refreshable {

        private final Frame ownerFrame;

        private final DefaultTableModel aggregateReportModel;
        private final JTable aggregateReportTable;
        private final JTextField aggStartDateField, aggEndDateField;

        private final DefaultTableModel atRiskReportModel;
        private final JTable atRiskReportTable;
        private final JTextField atRiskStartDateField, atRiskEndDateField, atRiskThresholdField;

        ReportsPanel(MainFrame frame) {
            this.ownerFrame = frame;
            setLayout(new BorderLayout(20, 20));
            setBorder(new EmptyBorder(25, 25, 25, 25));
            setBackground(AppStyles.BACKGROUND_COLOR);

            JLabel title = new JLabel("Attendance Reports");
            title.setFont(AppStyles.FONT_HEADER);
            add(title, BorderLayout.NORTH);

            JTabbedPane tabbedPane = new JTabbedPane();
            tabbedPane.setFont(AppStyles.FONT_BOLD);

            String todayStr = LocalDate.now().format(GLOBAL_DATE_FORMATTER);
            String firstDayOfMonthStr = LocalDate.now().withDayOfMonth(1).format(GLOBAL_DATE_FORMATTER);

            // --- Tab 1: Aggregate Report Panel ---
            JPanel aggregatePanel = new JPanel(new BorderLayout(15, 15));
            aggregatePanel.setOpaque(false);

            JPanel aggToolbar = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 0));
            aggToolbar.setOpaque(false);
            aggToolbar.add(new JLabel("Start Date:"));
            aggStartDateField = new JTextField(firstDayOfMonthStr, 10);
            aggToolbar.add(aggStartDateField);
            aggToolbar.add(new JLabel("End Date:"));
            aggEndDateField = new JTextField(todayStr, 10);
            aggToolbar.add(aggEndDateField);

            JButton aggGenerateButton = new JButton("Generate Report");
            aggGenerateButton.setBackground(AppStyles.PRIMARY_COLOR);
            aggGenerateButton.setForeground(Color.WHITE);
            aggGenerateButton.setFont(AppStyles.FONT_BOLD);
            aggGenerateButton.addActionListener(e -> generateAggregateReport());
            aggToolbar.add(aggGenerateButton);
            aggregatePanel.add(aggToolbar, BorderLayout.NORTH);

            aggregateReportModel = new DefaultTableModel(new String[]{"Student ID", "Name", "Total Present", "Total Absent", "Attendance %"}, 0) {
                @Override
                public boolean isCellEditable(int r, int c) {
                    return false;
                }
            };
            aggregateReportTable = CustomComponents.createModernTable(aggregateReportModel);
            aggregateReportTable.getColumnModel().getColumn(4).setCellRenderer(new CustomComponents.PercentageCellRenderer());
            aggregatePanel.add(new JScrollPane(aggregateReportTable), BorderLayout.CENTER);

            JPanel aggActionsPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
            aggActionsPanel.setOpaque(false);
            JButton aggExportButton = new JButton("Export to CSV");
            aggExportButton.setFont(AppStyles.FONT_BOLD);
            aggExportButton.addActionListener(e -> exportToCSV(aggregateReportTable, "AggregateReport"));
            aggActionsPanel.add(aggExportButton);
            aggregatePanel.add(aggActionsPanel, BorderLayout.SOUTH);

            tabbedPane.addTab("Aggregate Report", aggregatePanel);

            // --- Tab 2: Student At Risk Panel ---
            JPanel atRiskPanel = new JPanel(new BorderLayout(15, 15));
            atRiskPanel.setOpaque(false);

            JPanel atRiskToolbar = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 0));
            atRiskToolbar.setOpaque(false);
            atRiskToolbar.add(new JLabel("Start Date:"));
            atRiskStartDateField = new JTextField(firstDayOfMonthStr, 10);
            atRiskToolbar.add(atRiskStartDateField);
            atRiskToolbar.add(new JLabel("End Date:"));
            atRiskEndDateField = new JTextField(todayStr, 10);
            atRiskToolbar.add(atRiskEndDateField);
            atRiskToolbar.add(new JLabel("Threshold % (Below):"));
            atRiskThresholdField = new JTextField("75", 4);
            atRiskToolbar.add(atRiskThresholdField);

            JButton atRiskGenerateButton = new JButton("Find Students At Risk");
            atRiskGenerateButton.setBackground(AppStyles.RED);
            atRiskGenerateButton.setForeground(Color.WHITE);
            atRiskGenerateButton.setFont(AppStyles.FONT_BOLD);
            atRiskGenerateButton.addActionListener(e -> generateAtRiskReport());
            atRiskToolbar.add(atRiskGenerateButton);
            atRiskPanel.add(atRiskToolbar, BorderLayout.NORTH);

            atRiskReportModel = new DefaultTableModel(new String[]{"Student ID", "Name", "Total Present", "Total Absent", "Attendance %"}, 0) {
                @Override
                public boolean isCellEditable(int r, int c) {
                    return false;
                }
            };
            atRiskReportTable = CustomComponents.createModernTable(atRiskReportModel);
            atRiskReportTable.getColumnModel().getColumn(4).setCellRenderer(new CustomComponents.PercentageCellRenderer());
            atRiskPanel.add(new JScrollPane(atRiskReportTable), BorderLayout.CENTER);

            JPanel atRiskActionsPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
            atRiskActionsPanel.setOpaque(false);
            JButton atRiskExportButton = new JButton("Export to CSV");
            atRiskExportButton.setFont(AppStyles.FONT_BOLD);
            atRiskExportButton.addActionListener(e -> exportToCSV(atRiskReportTable, "AtRiskReport"));
            atRiskActionsPanel.add(atRiskExportButton);
            atRiskPanel.add(atRiskActionsPanel, BorderLayout.SOUTH);

            tabbedPane.addTab("Student At Risk Report", atRiskPanel);

            add(tabbedPane, BorderLayout.CENTER);
        }

        // Common calculation engine for both reports
        private Map<String, double[]> calculateReportData(String startDateStr, String endDateStr) throws DateTimeParseException, IllegalArgumentException {
            LocalDate startDate, endDate;
            startDate = LocalDate.parse(startDateStr, GLOBAL_DATE_FORMATTER);
            endDate = LocalDate.parse(endDateStr, GLOBAL_DATE_FORMATTER);

            if (startDate.isAfter(endDate)) {
                throw new IllegalArgumentException("Start Date cannot be after End Date.");
            }

            List<User> students = ExcelDataManager.getUsersByRole("Student");
            List<AttendanceRecord> allRecords = ExcelDataManager.getAllAttendance();

            List<AttendanceRecord> filteredRecords = allRecords.stream()
                    .filter(r -> {
                        try {
                            LocalDate recordDate = LocalDate.parse(r.date, GLOBAL_DATE_FORMATTER);
                            return !recordDate.isBefore(startDate) && !recordDate.isAfter(endDate);
                        } catch (DateTimeParseException e) {
                            return false;
                        }
                    })
                    .collect(Collectors.toList());

            Map<String, double[]> reportData = new HashMap<>(); // Stores {present, absent}
            students.forEach(student -> reportData.put(student.id, new double[]{0, 0}));

            for (AttendanceRecord record : filteredRecords) {
                double[] counts = reportData.get(record.studentId);
                if (counts != null) {
                    if ("Present".equals(record.status)) {
                        counts[0]++;
                    } else {
                        counts[1]++;
                    }
                }
            }
            return reportData;
        }

        // Logic for Tab 1
        private void generateAggregateReport() {
            Map<String, double[]> reportData;
            try {
                reportData = calculateReportData(aggStartDateField.getText(), aggEndDateField.getText());
            } catch (DateTimeParseException e) {
                JOptionPane.showMessageDialog(this, "Invalid date format. Please use yyyy-MM-dd.", "Date Error", JOptionPane.ERROR_MESSAGE);
                return;
            } catch (IllegalArgumentException e) {
                JOptionPane.showMessageDialog(this, e.getMessage(), "Date Error", JOptionPane.ERROR_MESSAGE);
                return;
            }

            aggregateReportModel.setRowCount(0);
            List<User> students = ExcelDataManager.getUsersByRole("Student");
            for (User student : students) {
                double[] counts = reportData.get(student.id);
                double present = counts[0];
                double absent = counts[1];
                double total = present + absent;
                double percentage = (total == 0) ? 100.0 : (present / total) * 100.0; // Default to 100 if no records

                aggregateReportModel.addRow(new Object[]{student.id, student.name, (int) present, (int) absent, percentage});
            }
        }

        // Logic for Tab 2 (WITH ERROR FIXED)
        private void generateAtRiskReport() {
            Map<String, double[]> reportData;
            double threshold;
            try {
                reportData = calculateReportData(atRiskStartDateField.getText(), atRiskEndDateField.getText());
                threshold = Double.parseDouble(atRiskThresholdField.getText());

                // --- BUG FIX IS HERE ---
                // The catch order is now correct: Specific (NumberFormat) before General (IllegalArgument)
            } catch (DateTimeParseException e) {
                JOptionPane.showMessageDialog(this, "Invalid date format. Please use yyyy-MM-dd.", "Date Error", JOptionPane.ERROR_MESSAGE);
                return;
            } catch (NumberFormatException e) {
                JOptionPane.showMessageDialog(this, "Threshold must be a number (e.g., 75).", "Input Error", JOptionPane.ERROR_MESSAGE);
                return;
            } catch (IllegalArgumentException e) {
                JOptionPane.showMessageDialog(this, e.getMessage(), "Date Error", JOptionPane.ERROR_MESSAGE);
                return;
            }
            // --- END OF FIX ---

            atRiskReportModel.setRowCount(0);
            List<User> students = ExcelDataManager.getUsersByRole("Student");
            for (User student : students) {
                double[] counts = reportData.get(student.id);
                double present = counts[0];
                double absent = counts[1];
                double total = present + absent;
                double percentage = (total == 0) ? 100.0 : (present / total) * 100.0; // Treat 0/0 as 100% (not at risk)

                // Only show students who have records AND are below the threshold
                if (total > 0 && percentage < threshold) {
                    atRiskReportModel.addRow(new Object[]{student.id, student.name, (int) present, (int) absent, percentage});
                }
            }
            if (atRiskReportModel.getRowCount() == 0) {
                JOptionPane.showMessageDialog(this, "No students are below the " + threshold + "% threshold for this period.", "Report Complete", JOptionPane.INFORMATION_MESSAGE);
            }
        }

        // Common Export logic
        private void exportToCSV(JTable table, String reportName) {
            TableModel model = table.getModel();
            if (model.getRowCount() == 0) {
                JOptionPane.showMessageDialog(this, "No data to export. Please generate a report first.", "Export Error", JOptionPane.WARNING_MESSAGE);
                return;
            }

            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setDialogTitle("Save Report as CSV");
            fileChooser.setSelectedFile(new File(reportName + "_" + LocalDate.now().format(GLOBAL_DATE_FORMATTER) + ".csv"));

            if (fileChooser.showSaveDialog(ownerFrame) == JFileChooser.APPROVE_OPTION) {
                File file = fileChooser.getSelectedFile();
                try (BufferedWriter bw = new BufferedWriter(new FileWriter(file))) {
                    for (int i = 0; i < model.getColumnCount(); i++) {
                        bw.write(model.getColumnName(i) + (i == model.getColumnCount() - 1 ? "" : ","));
                    }
                    bw.newLine();

                    for (int r = 0; r < model.getRowCount(); r++) {
                        for (int c = 0; c < model.getColumnCount(); c++) {
                            Object value = model.getValueAt(r, c);
                            if (value instanceof Double) {
                                value = String.format("%.2f", (Double) value);
                            }
                            bw.write("\"" + value.toString() + "\"" + (c == model.getColumnCount() - 1 ? "" : ","));
                        }
                        bw.newLine();
                    }
                    JOptionPane.showMessageDialog(ownerFrame, "Report exported successfully!", "Export Success", JOptionPane.INFORMATION_MESSAGE);
                } catch (IOException e) {
                    JOptionPane.showMessageDialog(ownerFrame, "Error exporting file: " + e.getMessage(), "Export Error", JOptionPane.ERROR_MESSAGE);
                }
            }
        }

        @Override
        public void refreshData() {
            aggregateReportModel.setRowCount(0);
            atRiskReportModel.setRowCount(0);
        }
    }

    // --- Add User Dialog ---
    static class AddUserDialog extends JDialog {

        AddUserDialog(Frame owner, String role, Runnable refreshCallback) {
            super(owner, "Add New " + role, true);

            boolean isStaff = role.equalsIgnoreCase("Staff");

            setSize(400, isStaff ? 300 : 250);
            setLocationRelativeTo(owner);
            setLayout(new GridBagLayout());

            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(8, 8, 8, 8);
            gbc.fill = GridBagConstraints.HORIZONTAL;
            gbc.gridx = 0;

            JTextField nameField = new JTextField();
            JPasswordField passField = new JPasswordField();

            gbc.gridy = 0;
            add(new JLabel("Full Name:"), gbc);
            gbc.gridx = 1;
            gbc.weightx = 1.0;
            add(nameField, gbc);

            gbc.gridy = 1;
            gbc.gridx = 0;
            gbc.weightx = 0;
            add(new JLabel("Password:"), gbc);
            gbc.gridx = 1;
            gbc.weightx = 1.0;
            add(passField, gbc);

            JComboBox<String> subjectSelector = new JComboBox<>(SUBJECT_LIST);
            if (isStaff) {
                gbc.gridy = 2;
                gbc.gridx = 0;
                gbc.weightx = 0;
                add(new JLabel("Assign Subject:"), gbc);
                gbc.gridx = 1;
                gbc.weightx = 1.0;
                add(subjectSelector, gbc);
            }

            JButton addButton = new JButton("Add User");
            addButton.setBackground(AppStyles.PRIMARY_COLOR);
            addButton.setForeground(Color.WHITE);
            addButton.setFont(AppStyles.FONT_BOLD);
            gbc.gridy = 3;
            gbc.gridx = 1;
            gbc.anchor = GridBagConstraints.EAST;
            gbc.fill = GridBagConstraints.NONE;
            add(addButton, gbc);

            addButton.addActionListener(e -> {
                String name = nameField.getText();
                String pass = new String(passField.getPassword());
                if (name.isEmpty() || pass.isEmpty()) {
                    JOptionPane.showMessageDialog(this, "Name and Password cannot be empty.", "Error", JOptionPane.ERROR_MESSAGE);
                    return;
                }

                String subject = "";
                if (isStaff) {
                    subject = (String) subjectSelector.getSelectedItem();
                }
                ExcelDataManager.addUser(name, pass, role, subject);
                refreshCallback.run();
                dispose();
            });
        }
    }

    // --- Excel Data Manager (NOW WITH PASSWORD UPDATE) ---
    static class ExcelDataManager {

        private static final String FILE_NAME = "college_data.xlsx";
        private static final String USERS_SHEET = "Users";
        private static final String ATTENDANCE_SHEET = "Attendance";

        public static void setupDatabase() {
            File dbFile = new File(FILE_NAME);
            Workbook workbook = null;
            try {
                if (!dbFile.exists()) {
                    workbook = new XSSFWorkbook();
                } else {
                    FileInputStream fis = new FileInputStream(dbFile);
                    workbook = new XSSFWorkbook(fis);
                    fis.close();
                }

                Sheet usersSheet = workbook.getSheet(USERS_SHEET);
                if (usersSheet == null) {
                    usersSheet = workbook.createSheet(USERS_SHEET);
                    Row header = usersSheet.createRow(0);
                    header.createCell(0).setCellValue("ID");
                    header.createCell(1).setCellValue("Password");
                    header.createCell(2).setCellValue("Name");
                    header.createCell(3).setCellValue("Role");
                    header.createCell(4).setCellValue("Subject");

                    Row adminRow = usersSheet.createRow(1);
                    adminRow.createCell(0).setCellValue("admin");
                    adminRow.createCell(1).setCellValue("admin123");
                    adminRow.createCell(2).setCellValue("Administrator");
                    adminRow.createCell(3).setCellValue("Admin");
                    adminRow.createCell(4).setCellValue("");
                } else {
                    Row header = usersSheet.getRow(0);
                    if (header == null) {
                        header = usersSheet.createRow(0);
                    }
                    if (header.getCell(4) == null) {
                        header.createCell(4).setCellValue("Subject");
                    }
                }

                Sheet attendanceSheet = workbook.getSheet(ATTENDANCE_SHEET);
                if (attendanceSheet == null) {
                    attendanceSheet = workbook.createSheet(ATTENDANCE_SHEET);
                    Row attHeader = attendanceSheet.createRow(0);
                    attHeader.createCell(0).setCellValue("StudentID");
                    attHeader.createCell(1).setCellValue("Date");
                    attHeader.createCell(2).setCellValue("Status");
                    attHeader.createCell(3).setCellValue("Subject");
                } else {
                    Row header = attendanceSheet.getRow(0);
                    if (header == null) {
                        header = attendanceSheet.createRow(0);
                    }
                    if (header.getCell(3) == null) {
                        header.createCell(3).setCellValue("Subject");
                    }
                }

                try (FileOutputStream fos = new FileOutputStream(FILE_NAME)) {
                    workbook.write(fos);
                }
                workbook.close();

            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        public static User authenticateUser(String username, String password) {
            try (FileInputStream fis = new FileInputStream(FILE_NAME); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheet(USERS_SHEET);
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) {
                        continue;
                    }

                    Cell nameCell = row.getCell(2);
                    Cell passCell = row.getCell(1);

                    if (nameCell != null && passCell != null
                            && nameCell.getStringCellValue().equals(username)
                            && passCell.getStringCellValue().equals(password)) {

                        String id = row.getCell(0).getStringCellValue();
                        String role = row.getCell(3).getStringCellValue();
                        String subject = "";

                        if (role.equals("Staff") && row.getCell(4) != null) {
                            subject = row.getCell(4).getStringCellValue();
                        }
                        return new User(id, password, username, role, subject);
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
            return null;
        }

        public static void addUser(String name, String password, String role, String subject) {
            String idPrefix = role.equalsIgnoreCase("Student") ? "STU" : "STAFF";

            List<User> usersOfRole = getUsersByRole(role);
            int lastIdNum = usersOfRole.stream()
                    .map(u -> u.id.replaceAll("[^0-9]", ""))
                    .filter(s -> !s.isEmpty())
                    .mapToInt(Integer::parseInt)
                    .max()
                    .orElse(0);

            String newId = String.format("%s%03d", idPrefix, lastIdNum + 1);

            try (FileInputStream fis = new FileInputStream(FILE_NAME); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheet(USERS_SHEET);
                Row newRow = sheet.createRow(sheet.getLastRowNum() + 1);
                newRow.createCell(0).setCellValue(newId);
                newRow.createCell(1).setCellValue(password);
                newRow.createCell(2).setCellValue(name);
                newRow.createCell(3).setCellValue(role);
                newRow.createCell(4).setCellValue(subject != null ? subject : "");

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
                    Row foundRow = sheet.getRow(rowToRemove);
                    if (foundRow != null) {
                        sheet.removeRow(foundRow);
                    }
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

        // NEW: Method to update a user's password
        public static void updatePassword(String userId, String newPassword) {
            File file = new File(FILE_NAME);
            try (FileInputStream fis = new FileInputStream(file); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheet(USERS_SHEET);
                boolean updated = false;
                for (Row row : sheet) {
                    if (row.getRowNum() == 0 || row.getCell(0) == null) {
                        continue;
                    }

                    if (row.getCell(0).getStringCellValue().equalsIgnoreCase(userId)) {
                        Cell passCell = row.getCell(1);
                        if (passCell == null) {
                            passCell = row.createCell(1);
                        }
                        passCell.setCellValue(newPassword);
                        updated = true;
                        break;
                    }
                }

                if (updated) {
                    try (FileOutputStream fos = new FileOutputStream(file)) {
                        workbook.write(fos);
                    }
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
                        String id = row.getCell(0).getStringCellValue();
                        String name = row.getCell(2).getStringCellValue();
                        String subject = "";
                        if (role.equals("Staff") && row.getCell(4) != null) {
                            subject = row.getCell(4).getStringCellValue();
                        }
                        users.add(new User(id, "", name, role, subject));
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
                        String name = row.getCell(2).getStringCellValue();
                        String role = row.getCell(3).getStringCellValue();
                        String subject = "";
                        if (role.equals("Staff") && row.getCell(4) != null) {
                            subject = row.getCell(4).getStringCellValue();
                        }
                        return new User(userId, "", name, role, subject);
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
                if (sheet == null) {
                    return records;
                }

                for (Row row : sheet) {
                    if (row.getRowNum() == 0 || row.getCell(0) == null) {
                        continue;
                    }

                    String subject = "General";
                    Cell subjectCell = row.getCell(3);
                    if (subjectCell != null && !subjectCell.getStringCellValue().isEmpty()) {
                        subject = subjectCell.getStringCellValue();
                    }

                    records.add(new AttendanceRecord(
                            row.getCell(0).getStringCellValue(),
                            row.getCell(1).getStringCellValue(),
                            row.getCell(2).getStringCellValue(),
                            subject
                    ));
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
            return records;
        }

        public static List<AttendanceRecord> getAllAttendanceForToday() {
            String today = LocalDate.now().format(GLOBAL_DATE_FORMATTER);
            return getAllAttendance().stream()
                    .filter(r -> r.date.equals(today))
                    .collect(Collectors.toList());
        }

        public static Map<String, Double> getOverallAttendanceForToday() {
            List<AttendanceRecord> todayRecords = getAllAttendanceForToday();
            return calculateStatsFromRecords(todayRecords);
        }

        public static Map<String, Double> getOverallAttendanceForToday(String subject) {
            List<AttendanceRecord> todayRecords = getAllAttendanceForToday().stream()
                    .filter(r -> r.subject.equals(subject))
                    .collect(Collectors.toList());
            return calculateStatsFromRecords(todayRecords);
        }

        private static Map<String, Double> calculateStatsFromRecords(List<AttendanceRecord> records) {
            double present = 0;
            double absent = 0;
            for (AttendanceRecord r : records) {
                if ("Present".equals(r.status)) {
                    present++; 
                }else {
                    absent++;
                }
            }

            double total = present + absent;
            double percentage = (total == 0) ? 100.0 : (present / total) * 100.0;

            Map<String, Double> stats = new HashMap<>();
            stats.put("present", present);
            stats.put("absent", absent);
            stats.put("total", total);
            stats.put("percentage", percentage);
            return stats;
        }

        public static List<AttendanceRecord> getAttendanceForStudent(String studentId) {
            return getAllAttendance().stream()
                    .filter(r -> r.studentId.equalsIgnoreCase(studentId))
                    .sorted((r1, r2) -> r2.date.compareTo(r1.date)) // Sort by date, most recent first
                    .collect(Collectors.toList());
        }

        public static boolean hasAttendanceBeenMarked(String dateStr, String subject) {
            return getAllAttendance().stream()
                    .anyMatch(r -> r.date.equals(dateStr) && r.subject.equals(subject));
        }

        public static String getStudentStatusForDate(String studentId, String dateStr, String subject) {
            return getAllAttendance().stream()
                    .filter(r -> r.date.equals(dateStr) && r.studentId.equals(studentId) && r.subject.equals(subject))
                    .map(r -> r.status)
                    .findFirst().orElse("Absent");
        }

        public static void markAttendance(List<AttendanceRecord> records, String dateStr, String subject) {
            try (FileInputStream fis = new FileInputStream(FILE_NAME); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheet(ATTENDANCE_SHEET);

                List<Integer> rowsToRemove = new ArrayList<>();
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) {
                        continue;
                    }

                    Cell dateCell = row.getCell(1);
                    Cell subjectCell = row.getCell(3);
                    String rowDate = (dateCell != null) ? dateCell.getStringCellValue() : "";
                    String rowSubject = (subjectCell != null && !subjectCell.getStringCellValue().isEmpty()) ? subjectCell.getStringCellValue() : "General";

                    if (rowDate.equals(dateStr) && rowSubject.equals(subject)) {
                        rowsToRemove.add(row.getRowNum());
                    }
                }

                rowsToRemove.sort(Comparator.reverseOrder());
                int removedCount = 0;
                for (int rowIndex : rowsToRemove) {
                    Row row = sheet.getRow(rowIndex);
                    if (row != null) {
                        sheet.removeRow(row);
                        removedCount++;
                    }
                }

                if (removedCount > 0) {
                    int firstRowRemoved = rowsToRemove.get(rowsToRemove.size() - 1);
                    int lastRowInSheet = sheet.getLastRowNum();
                    if (firstRowRemoved <= lastRowInSheet) {
                        sheet.shiftRows(firstRowRemoved + 1, lastRowInSheet + removedCount, -removedCount);
                    }
                }

                int lastRow = sheet.getLastRowNum();
                for (int i = 0; i < records.size(); i++) {
                    AttendanceRecord record = records.get(i);
                    Row newRow = sheet.createRow(lastRow + 1 + i);
                    newRow.createCell(0).setCellValue(record.studentId);
                    newRow.createCell(1).setCellValue(record.date);
                    newRow.createCell(2).setCellValue(record.status);
                    newRow.createCell(3).setCellValue(record.subject);
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
