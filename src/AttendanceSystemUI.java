
import com.formdev.flatlaf.themes.FlatMacLightLaf;

// JCalendar (com.toedter) imports are now REMOVED
import javax.swing.*;
import javax.swing.border.Border;
import javax.swing.border.EmptyBorder;
import javax.swing.event.TableModelEvent;
import javax.swing.event.TableModelListener;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellEditor;
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
import java.awt.Point;
import java.awt.RenderingHints;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener; // Specific import
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.DayOfWeek; // NEW import for calendar logic
import java.time.LocalDate; // NEW import for calendar logic
import java.time.YearMonth; // NEW import for calendar logic
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.time.format.TextStyle; // NEW import for calendar logic
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale; // NEW import for calendar logic
import java.util.Map;
import java.util.Objects;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AttendanceSystemUI {

    public static final String[] SUBJECT_LIST = {"General", "Math", "Physics", "Chemistry", "History", "English", "Biology"};
    private static final DateTimeFormatter GLOBAL_DATE_FORMATTER = DateTimeFormatter.ISO_LOCAL_DATE;

    // --- Main Entry Point ---
    public static void main(String[] args) {
        FlatMacLightLaf.setup();

        UIManager.put("Button.arc", 999);
        UIManager.put("Component.arc", 12);
        UIManager.put("ProgressBar.arc", 12);
        UIManager.put("TextComponent.arc", 8);
        UIManager.put("Table.showVerticalLines", false);
        UIManager.put("Table.intercellSpacing", new Dimension(0, 1));
        UIManager.put("TableHeader.height", 40);
        UIManager.put("TableHeader.background", new Color(245, 248, 251));
        UIManager.put("TableHeader.separatorColor", new Color(230, 230, 230));
        UIManager.put("Component.arrowType", "triangle");
        UIManager.put("Component.focusWidth", 1);
        UIManager.put("Component.focusColor", AppStyles.PRIMARY_COLOR);

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
        public static final Color ORANGE = new Color(243, 156, 18);
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
            this(id, password, name, role, "");
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

            switch (currentUser.role) {
                case "Admin":
                case "Staff":
                    sideNavPanel.setActive("Dashboard");
                    showPanel("HomeDashboard");
                    break;
                case "Student":
                    sideNavPanel.setActive("Dashboard");
                    showPanel("StudentDashboard");
                    break;
            }
        }

        private void setupContentPanels() {
            contentPanel.add(new HomeDashboardPanel(this), "HomeDashboard");
            contentPanel.add(new StaffAttendancePanel(this), "Staff");
            contentPanel.add(new AdminPanel(this), "Admin");
            contentPanel.add(new StudentDashboardPanel(this), "StudentDashboard");

            if ("Admin".equals(currentUser.role)) {
                contentPanel.add(new AdminReportsPanel(this), "Reports");
            } else if ("Staff".equals(currentUser.role)) {
                contentPanel.add(new StaffReportsPanel(this), "Reports");
            }
        }

        public void showPanel(String panelName) {
            cardLayout.show(contentPanel, panelName);
            for (Component component : contentPanel.getComponents()) {
                String compName = component.getClass().getSimpleName();
                if ((compName.equals(panelName) || compName.equals(panelName + "Panel")) && component.isVisible()) {
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
                    addNavButton("🏠", "Dashboard", "HomeDashboard");
                    addNavButton("👤", "Students", "Admin");
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

    // --- Reusable UI Components ---
    static class CustomComponents {

        public static JPanel createCardPanel() {
            JPanel card = new JPanel();
            card.setBackground(Color.WHITE);
            card.setBorder(BorderFactory.createCompoundBorder(
                    BorderFactory.createLineBorder(AppStyles.BORDER_COLOR, 1, true),
                    new EmptyBorder(15, 15, 15, 15)));
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
            JPanel textPanel = new JPanel(new GridLayout(2, 1));
            textPanel.setOpaque(false);
            JLabel valueLabel = new JLabel(value);
            valueLabel.setFont(new Font("Inter", Font.BOLD, 26));
            valueLabel.setForeground(AppStyles.PRIMARY_TEXT_COLOR);
            JLabel titleLabel = new JLabel(title);
            titleLabel.setFont(AppStyles.FONT_SMALL);
            titleLabel.setForeground(AppStyles.SECONDARY_TEXT_COLOR);
            textPanel.add(valueLabel);
            textPanel.add(titleLabel);
            card.add(textPanel, BorderLayout.CENTER);
            return card;
        }

        public static class StatusCellRenderer extends DefaultTableCellRenderer {

            @Override
            public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
                super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
                if (value == null) {
                    return this;
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

    // --- Admin/Staff Home Dashboard Panel ---
    static class HomeDashboardPanel extends JPanel implements Refreshable {

        private final User currentUser;
        private final JPanel statsPanel = new JPanel(new GridLayout(1, 4, 20, 20));

        HomeDashboardPanel(MainFrame frame) {
            this.currentUser = frame.getCurrentUser();
            setLayout(new BorderLayout(20, 20));
            setBorder(new EmptyBorder(25, 25, 25, 25));
            setBackground(AppStyles.BACKGROUND_COLOR);
            JLabel titleLabel = new JLabel("Welcome, " + currentUser.name);
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
                statsPanel.add(CustomComponents.createStatCard("Total Staff", String.valueOf(staff.size()), "👤", AppStyles.ORANGE));
                statsPanel.add(CustomComponents.createStatCard("Total Lectures Today", String.valueOf(todayStats.get("total").intValue()), "📚", AppStyles.GREEN));
                statsPanel.add(CustomComponents.createStatCard("Overall Absent Today", String.valueOf(todayStats.get("absent").intValue()), "❌", AppStyles.RED));
            } else if ("Staff".equals(currentUser.role)) {
                Map<String, Double> subjectStats = ExcelDataManager.getOverallAttendanceForToday(currentUser.subject);
                statsPanel.add(CustomComponents.createStatCard("My Assigned Subject", currentUser.subject, "📚", AppStyles.PRIMARY_COLOR));
                statsPanel.add(CustomComponents.createStatCard("Att. % (My Subject)", String.format("%.0f%%", subjectStats.get("percentage")), "📈", AppStyles.GREEN));
                statsPanel.add(CustomComponents.createStatCard("Total Lectures (My Subject)", String.valueOf(subjectStats.get("total").intValue()), "✍️", AppStyles.ORANGE));
                statsPanel.add(CustomComponents.createStatCard("Absent (My Subject)", String.valueOf(subjectStats.get("absent").intValue()), "❌", AppStyles.RED));
            }
            statsPanel.revalidate();
            statsPanel.repaint();
        }
    }

    // --- Staff View Panel (WITH NEW CALENDAR) ---
    static class StaffAttendancePanel extends JPanel implements Refreshable {

        private final DefaultTableModel tableModel;
        private final JPanel statsPanel = new JPanel(new GridLayout(1, 4, 20, 20));
        private final ModernDatePickerButton datePickerButton; // NEW
        private final String assignedSubject;

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

            datePickerButton = new ModernDatePickerButton(frame); // NEW
            toolbar.add(datePickerButton);

            JLabel subjectLabel = new JLabel("Subject: " + this.assignedSubject);
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
                    int row = table.rowAtPoint(e.getPoint()), col = table.columnAtPoint(e.getPoint());
                    if (row >= 0 && col == 2) {
                        String newStatus = "Present".equals(tableModel.getValueAt(row, 2)) ? "Absent" : "Present";
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

        private String getSelectedDateString() {
            LocalDate selectedDate = datePickerButton.getSelectedDate();
            if (selectedDate == null) {
                JOptionPane.showMessageDialog(this, "Please select a valid date.", "Date Error", JOptionPane.ERROR_MESSAGE);
                return null;
            }
            return selectedDate.format(GLOBAL_DATE_FORMATTER);
        }

        private void loadRosterForAssignedSubject() {
            String dateStr = getSelectedDateString();
            if (dateStr == null) {
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
            datePickerButton.setSelectedDate(LocalDate.now()); // Set calendar to today
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
            statsPanel.add(CustomComponents.createStatCard("Total Roster", String.valueOf(total), "🎓", AppStyles.ORANGE));
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
            String dateStr = getSelectedDateString();
            if (dateStr == null) {
                return;
            }
            List<AttendanceRecord> records = new ArrayList<>();
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                records.add(new AttendanceRecord((String) tableModel.getValueAt(i, 0), dateStr, (String) tableModel.getValueAt(i, 2), this.assignedSubject));
            }
            if (ExcelDataManager.hasAttendanceBeenMarked(dateStr, this.assignedSubject)) {
                int choice = JOptionPane.showConfirmDialog(this, "Overwrite attendance for '" + this.assignedSubject + "' on " + dateStr + "?", "Confirm Overwrite", JOptionPane.YES_NO_OPTION);
                if (choice == JOptionPane.NO_OPTION) {
                    return;
                }
            }
            ExcelDataManager.markAttendance(records, dateStr, this.assignedSubject);
            JOptionPane.showMessageDialog(this, "Attendance for " + this.assignedSubject + " on " + dateStr + " saved successfully!", "Success", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    // --- Admin Panel (WITH EDITABLE STATUS COLUMN) ---
    static class AdminPanel extends JPanel implements Refreshable {

        private final JTable studentTable, staffTable, reportTable;
        private final DefaultTableModel studentModel, staffModel, reportModel;
        private boolean isUpdatingByCode = false; // NEW: Flag to prevent event loops

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

            reportModel = new DefaultTableModel(new String[]{"Student ID", "Name", "Date", "Subject", "Status"}, 0) {
                @Override
                public boolean isCellEditable(int row, int column) {
                    return column == 4;
                } // Only Status (col 4) is editable
            };
            reportTable = CustomComponents.createModernTable(reportModel);

            JComboBox<String> statusComboBox = new JComboBox<>(new String[]{"Present", "Absent"});
            statusComboBox.setFont(AppStyles.FONT_NORMAL);
            DefaultCellEditor statusEditor = new DefaultCellEditor(statusComboBox);
            reportTable.getColumnModel().getColumn(4).setCellEditor(statusEditor);
            reportTable.getColumnModel().getColumn(4).setCellRenderer(new CustomComponents.StatusCellRenderer());

            reportModel.addTableModelListener(e -> {
                if (isUpdatingByCode || e.getType() != TableModelEvent.UPDATE) {
                    return; // Prevent loop
                }
                int row = e.getFirstRow(), col = e.getColumn();
                if (col == 4) { // Status column was changed
                    String studentId = (String) reportModel.getValueAt(row, 0);
                    String date = (String) reportModel.getValueAt(row, 2);
                    String subject = (String) reportModel.getValueAt(row, 3);
                    String newStatus = (String) reportModel.getValueAt(row, 4);

                    new Thread(() -> { // Run Excel save on a background thread
                        ExcelDataManager.updateSingleAttendanceRecord(studentId, date, subject, newStatus);
                        SwingUtilities.invokeLater(() -> System.out.println("Updated record for " + studentId + " on " + date));
                    }).start();
                }
            });

            tabbedPane.addTab("Raw Attendance Data (Editable Status)", new JScrollPane(reportTable));
            add(tabbedPane, BorderLayout.CENTER);
            refreshData();
        }

        private JPanel createManagementPanel(JTable table, String role, Frame owner, Runnable refreshCallback) {
            JPanel panel = new JPanel(new BorderLayout(10, 10));
            panel.add(new JScrollPane(table), BorderLayout.CENTER);
            JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
            JButton addButton = new JButton("➕ Add New " + role);
            JButton removeButton = new JButton("❌ Remove Selected");
            JButton setPasswordButton = new JButton("🔑 Set Password");
            addButton.setFont(AppStyles.FONT_BOLD);
            removeButton.setFont(AppStyles.FONT_BOLD);
            setPasswordButton.setFont(AppStyles.FONT_BOLD);
            addButton.setBackground(AppStyles.GREEN);
            addButton.setForeground(Color.WHITE);
            setPasswordButton.setBackground(AppStyles.ORANGE);
            addButton.addActionListener(e -> new AddUserDialog(owner, role, refreshCallback).setVisible(true));
            removeButton.addActionListener(e -> {
                int selectedRow = table.getSelectedRow();
                if (selectedRow != -1) {
                    String id = (String) table.getModel().getValueAt(table.convertRowIndexToModel(selectedRow), 0);
                    String name = (String) table.getModel().getValueAt(table.convertRowIndexToModel(selectedRow), 1);
                    int confirm = JOptionPane.showConfirmDialog(this, "Remove user: " + name + " (ID: " + id + ")?", "Confirm Deletion", JOptionPane.YES_NO_OPTION, JOptionPane.WARNING_MESSAGE);
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
            isUpdatingByCode = true; // Set flag
            studentModel.setRowCount(0);
            ExcelDataManager.getUsersByRole("Student").forEach(u -> studentModel.addRow(new Object[]{u.id, u.name, u.role}));
            staffModel.setRowCount(0);
            ExcelDataManager.getUsersByRole("Staff").forEach(u -> staffModel.addRow(new Object[]{u.id, u.name, u.role, u.subject}));
            reportModel.setRowCount(0);
            ExcelDataManager.getAllAttendance().forEach(r -> {
                User student = ExcelDataManager.getUserById(r.studentId);
                reportModel.addRow(new Object[]{r.studentId, student != null ? student.name : "N/A", r.date, r.subject, r.status});
            });
            isUpdatingByCode = false; // Release flag
        }
    }

    // --- STUDENT DASHBOARD (NEW LAYOUT) ---
    static class StudentDashboardPanel extends JPanel implements Refreshable {

        private final DefaultTableModel historyTableModel, subjectSummaryModel;
        private final JPanel statsPanel = new JPanel(new GridLayout(1, 4, 20, 20));
        private final MainFrame mainFrame;
        private final JLabel lowAttendanceWarning;
        private final JComboBox<String> historyFilterDropdown;
        private List<AttendanceRecord> allStudentRecords = new ArrayList<>();

        StudentDashboardPanel(MainFrame frame) {
            this.mainFrame = frame;
            setLayout(new BorderLayout(20, 20));
            setBorder(new EmptyBorder(25, 25, 25, 25));
            setBackground(AppStyles.BACKGROUND_COLOR);

            JPanel headerPanel = new JPanel(new BorderLayout());
            headerPanel.setOpaque(false);
            JLabel title = new JLabel("Welcome, " + frame.getCurrentUser().name);
            title.setFont(AppStyles.FONT_HEADER);
            headerPanel.add(title, BorderLayout.NORTH);
            lowAttendanceWarning = new JLabel("⚠️ Your attendance is below 75% in one or more subjects. Please review your Subject Summary.");
            lowAttendanceWarning.setFont(AppStyles.FONT_BOLD);
            lowAttendanceWarning.setForeground(Color.WHITE);
            lowAttendanceWarning.setBackground(AppStyles.ORANGE.darker());
            lowAttendanceWarning.setOpaque(true);
            lowAttendanceWarning.setBorder(new EmptyBorder(10, 15, 10, 15));
            lowAttendanceWarning.setVisible(false);
            headerPanel.add(lowAttendanceWarning, BorderLayout.SOUTH);
            add(headerPanel, BorderLayout.NORTH);
            statsPanel.setOpaque(false);

            // --- Subject Summary Table ---
            subjectSummaryModel = new DefaultTableModel(new String[]{"Subject", "Total Lectures", "Present", "Absent", "Percentage"}, 0) {
                @Override
                public boolean isCellEditable(int r, int c) {
                    return false;
                }
            };
            JTable subjectTable = CustomComponents.createModernTable(subjectSummaryModel);
            subjectTable.getColumnModel().getColumn(4).setCellRenderer(new CustomComponents.PercentageCellRenderer());
            JScrollPane subjectScrollPane = new JScrollPane(subjectTable);

            JLabel summaryTitle = new JLabel("Overall Subject Summary");
            summaryTitle.setFont(AppStyles.FONT_HEADER.deriveFont(20f));
            summaryTitle.setBorder(new EmptyBorder(10, 5, 10, 5));

            JPanel subjectSummaryPanel = new JPanel(new BorderLayout(0, 10));
            subjectSummaryPanel.setOpaque(false);
            subjectSummaryPanel.add(summaryTitle, BorderLayout.NORTH);
            subjectSummaryPanel.add(subjectScrollPane, BorderLayout.CENTER);

            // --- History Tabbed Pane ---
            JTabbedPane tabbedPane = new JTabbedPane();
            tabbedPane.setFont(AppStyles.FONT_BOLD);
            JPanel historyPanel = new JPanel(new BorderLayout(10, 10));
            historyPanel.setOpaque(false);
            JPanel historyToolbar = new JPanel(new FlowLayout(FlowLayout.LEFT));
            historyToolbar.setOpaque(false);
            historyToolbar.add(new JLabel("Filter by Subject:"));
            DefaultComboBoxModel<String> filterModel = new DefaultComboBoxModel<>();
            filterModel.addElement("All Subjects");
            for (String s : SUBJECT_LIST) {
                filterModel.addElement(s);
            }
            historyFilterDropdown = new JComboBox<>(filterModel);
            historyFilterDropdown.addActionListener(e -> populateHistoryTable());
            historyToolbar.add(historyFilterDropdown);
            historyPanel.add(historyToolbar, BorderLayout.NORTH);

            historyTableModel = new DefaultTableModel(new String[]{"Date", "Subject", "Status"}, 0) {
                @Override
                public boolean isCellEditable(int r, int c) {
                    return false;
                }
            };
            JTable historyTable = CustomComponents.createModernTable(historyTableModel);
            historyTable.getColumnModel().getColumn(2).setCellRenderer(new CustomComponents.StatusCellRenderer());
            historyPanel.add(new JScrollPane(historyTable), BorderLayout.CENTER);
            tabbedPane.addTab("Recent History Log", historyPanel);

            // --- Main Center Panel Layout ---
            JPanel centerPanel = new JPanel(new BorderLayout(20, 20));
            centerPanel.setOpaque(false);
            centerPanel.add(statsPanel, BorderLayout.NORTH); // Stats on top
            centerPanel.add(subjectSummaryPanel, BorderLayout.CENTER); // Summary table in middle
            centerPanel.add(tabbedPane, BorderLayout.SOUTH); // History tabs at bottom
            add(centerPanel, BorderLayout.CENTER);

            refreshData();
        }

        private void populateHistoryTable() {
            historyTableModel.setRowCount(0);
            String filter = (String) historyFilterDropdown.getSelectedItem();
            boolean showAll = "All Subjects".equals(filter);
            for (AttendanceRecord r : allStudentRecords) {
                if (showAll || r.subject.equals(filter)) {
                    historyTableModel.addRow(new Object[]{r.date, r.subject, r.status});
                }
            }
        }

        public void refreshData(User currentUser) {
            subjectSummaryModel.setRowCount(0);
            allStudentRecords = ExcelDataManager.getAttendanceForStudent(currentUser.id);
            populateHistoryTable();

            int overallPresent = 0, overallAbsent = 0;
            boolean atRiskFlag = false;
            Map<String, int[]> subjectSummaryData = new LinkedHashMap<>();
            for (AttendanceRecord r : allStudentRecords) {
                if ("Present".equals(r.status)) {
                    overallPresent++;
                } else {
                    overallAbsent++;
                }
                int[] counts = subjectSummaryData.computeIfAbsent(r.subject, k -> new int[2]);
                if ("Present".equals(r.status)) {
                    counts[0]++;
                } else {
                    counts[1]++;
                }
            }
            int overallTotal = overallPresent + overallAbsent;
            double overallPercentage = (overallTotal == 0) ? 100.0 : ((double) overallPresent / overallTotal) * 100.0;
            if (overallPercentage < 75.0 && overallTotal > 0) {
                atRiskFlag = true;
            }

            statsPanel.removeAll();
            statsPanel.add(CustomComponents.createStatCard("Total Lectures", String.valueOf(overallTotal), "📚", AppStyles.ORANGE));
            statsPanel.add(CustomComponents.createStatCard("Total Present", String.valueOf(overallPresent), "✅", AppStyles.GREEN));
            statsPanel.add(CustomComponents.createStatCard("Total Absent", String.valueOf(overallAbsent), "❌", AppStyles.RED));
            statsPanel.add(CustomComponents.createStatCard("Overall Percentage", String.format("%.0f%%", overallPercentage), "📈", AppStyles.PRIMARY_COLOR));

            for (Map.Entry<String, int[]> entry : subjectSummaryData.entrySet()) {
                String subject = entry.getKey();
                int present = entry.getValue()[0], absent = entry.getValue()[1], total = present + absent;
                double percentage = (total == 0) ? 100.0 : ((double) present / total) * 100.0;
                if (percentage < 75.0 && total > 0) {
                    atRiskFlag = true;
                }
                subjectSummaryModel.addRow(new Object[]{subject, total, present, absent, percentage});
            }
            lowAttendanceWarning.setVisible(atRiskFlag);
            statsPanel.revalidate();
            statsPanel.repaint();
        }

        @Override
        public void refreshData() {
            refreshData(mainFrame.getCurrentUser());
        }
    }

    // --- BASE Reports Panel ---
    static abstract class BaseReportsPanel extends JPanel implements Refreshable {

        protected final Frame ownerFrame;
        protected final DefaultTableModel aggregateReportModel, atRiskReportModel;
        protected final JTable aggregateReportTable, atRiskReportTable;
        protected final ModernDatePickerButton aggStartDateChooser, aggEndDateChooser; // NEW
        protected final ModernDatePickerButton atRiskStartDateChooser, atRiskEndDateChooser; // NEW
        protected final JTextField atRiskThresholdField;

        public BaseReportsPanel(MainFrame frame) {
            this.ownerFrame = frame;
            setLayout(new BorderLayout(20, 20));
            setBorder(new EmptyBorder(25, 25, 25, 25));
            setBackground(AppStyles.BACKGROUND_COLOR);
            JLabel title = new JLabel("Attendance Reports");
            title.setFont(AppStyles.FONT_HEADER);
            add(title, BorderLayout.NORTH);
            JTabbedPane tabbedPane = new JTabbedPane();
            tabbedPane.setFont(AppStyles.FONT_BOLD);

            aggregateReportModel = createReportModel();
            aggregateReportTable = CustomComponents.createModernTable(aggregateReportModel);
            aggregateReportTable.getColumnModel().getColumn(5).setCellRenderer(new CustomComponents.PercentageCellRenderer()); // Col 5 is %
            atRiskReportModel = createReportModel();
            atRiskReportTable = CustomComponents.createModernTable(atRiskReportModel);
            atRiskReportTable.getColumnModel().getColumn(5).setCellRenderer(new CustomComponents.PercentageCellRenderer()); // Col 5 is %

            aggStartDateChooser = new ModernDatePickerButton(frame);
            aggEndDateChooser = new ModernDatePickerButton(frame);
            atRiskStartDateChooser = new ModernDatePickerButton(frame);
            atRiskEndDateChooser = new ModernDatePickerButton(frame);

            LocalDate firstDayOfMonth = LocalDate.now().withDayOfMonth(1);
            aggStartDateChooser.setSelectedDate(firstDayOfMonth);
            atRiskStartDateChooser.setSelectedDate(firstDayOfMonth);

            atRiskThresholdField = new JTextField("75", 4);
            atRiskThresholdField.setFont(AppStyles.FONT_NORMAL);
            atRiskThresholdField.setPreferredSize(new Dimension(50, 40));

            tabbedPane.addTab("Aggregate Report", createAggregateReportTab());
            tabbedPane.addTab("Student At Risk Report", createAtRiskReportTab());
            add(tabbedPane, BorderLayout.CENTER);
        }

        // REFACTORED: Create separate abstract methods for each toolbar
        protected abstract JPanel createAggToolbarExtras();

        protected abstract JPanel createAtRiskToolbarExtras();

        protected abstract String getAggSubjectFilter();

        protected abstract String getAtRiskSubjectFilter();

        private JPanel createAggregateReportTab() {
            JPanel aggregatePanel = new JPanel(new BorderLayout(15, 15));
            aggregatePanel.setOpaque(false);
            JPanel aggToolbar = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 0));
            aggToolbar.setOpaque(false);
            aggToolbar.add(new JLabel("Start Date:"));
            aggToolbar.add(aggStartDateChooser);
            aggToolbar.add(new JLabel("End Date:"));
            aggToolbar.add(aggEndDateChooser);
            aggToolbar.add(createAggToolbarExtras()); // Call specific method
            JButton aggGenerateButton = new JButton("Generate Report");
            aggGenerateButton.setBackground(AppStyles.PRIMARY_COLOR);
            aggGenerateButton.setForeground(Color.WHITE);
            aggGenerateButton.setFont(AppStyles.FONT_BOLD);
            aggGenerateButton.addActionListener(e -> generateAggregateReport());
            aggToolbar.add(aggGenerateButton);
            aggregatePanel.add(aggToolbar, BorderLayout.NORTH);
            aggregatePanel.add(new JScrollPane(aggregateReportTable), BorderLayout.CENTER);
            JPanel aggActionsPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
            aggActionsPanel.setOpaque(false);
            JButton aggExportButton = new JButton("Export to CSV");
            aggExportButton.setFont(AppStyles.FONT_BOLD);
            aggExportButton.addActionListener(e -> exportToCSV(aggregateReportTable, "AggregateReport"));
            aggActionsPanel.add(aggExportButton);
            aggregatePanel.add(aggActionsPanel, BorderLayout.SOUTH);
            return aggregatePanel;
        }

        private JPanel createAtRiskReportTab() {
            JPanel atRiskPanel = new JPanel(new BorderLayout(15, 15));
            atRiskPanel.setOpaque(false);
            JPanel atRiskToolbar = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 0));
            atRiskToolbar.setOpaque(false);
            atRiskToolbar.add(new JLabel("Start Date:"));
            atRiskToolbar.add(atRiskStartDateChooser);
            atRiskToolbar.add(new JLabel("End Date:"));
            atRiskToolbar.add(atRiskEndDateChooser);
            atRiskToolbar.add(createAtRiskToolbarExtras()); // Call specific method
            atRiskToolbar.add(new JLabel("Threshold % (Below):"));
            atRiskToolbar.add(atRiskThresholdField);
            JButton atRiskGenerateButton = new JButton("Find Students At Risk");
            atRiskGenerateButton.setBackground(AppStyles.RED);
            atRiskGenerateButton.setForeground(Color.WHITE);
            atRiskGenerateButton.setFont(AppStyles.FONT_BOLD);
            atRiskGenerateButton.addActionListener(e -> generateAtRiskReport());
            atRiskToolbar.add(atRiskGenerateButton);
            atRiskPanel.add(atRiskToolbar, BorderLayout.NORTH);
            atRiskPanel.add(new JScrollPane(atRiskReportTable), BorderLayout.CENTER);
            JPanel atRiskActionsPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
            atRiskActionsPanel.setOpaque(false);
            JButton atRiskExportButton = new JButton("Export to CSV");
            atRiskExportButton.setFont(AppStyles.FONT_BOLD);
            atRiskExportButton.addActionListener(e -> exportToCSV(atRiskReportTable, "AtRiskReport"));
            atRiskActionsPanel.add(atRiskExportButton);
            atRiskPanel.add(atRiskActionsPanel, BorderLayout.SOUTH);
            return atRiskPanel;
        }

        protected Map<String, double[]> calculateReportData(LocalDate start, LocalDate end, String subjectFilter) throws IllegalArgumentException {
            if (start == null || end == null) {
                throw new IllegalArgumentException("Dates cannot be empty.");
            }
            if (start.isAfter(end)) {
                throw new IllegalArgumentException("Start Date cannot be after End Date.");
            }

            List<User> students = ExcelDataManager.getUsersByRole("Student");
            List<AttendanceRecord> allRecords = ExcelDataManager.getAllAttendance();
            boolean allSubjects = "All Subjects".equals(subjectFilter);

            List<AttendanceRecord> subjectFilteredRecords = allRecords.stream()
                    .filter(r -> allSubjects || r.subject.equals(subjectFilter))
                    .collect(Collectors.toList());

            List<AttendanceRecord> dateFilteredRecords = subjectFilteredRecords.stream()
                    .filter(r -> {
                        try {
                            LocalDate recordDate = LocalDate.parse(r.date, GLOBAL_DATE_FORMATTER);
                            return !recordDate.isBefore(start) && !recordDate.isAfter(end);
                        } catch (DateTimeParseException e) {
                            return false;
                        }
                    })
                    .collect(Collectors.toList());

            Map<String, double[]> reportData = new HashMap<>(); // Stores {present, absent}
            students.forEach(student -> reportData.put(student.id, new double[]{0, 0}));
            for (AttendanceRecord record : dateFilteredRecords) {
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

        private DefaultTableModel createReportModel() {
            return new DefaultTableModel(new String[]{"Student ID", "Name", "Total Lectures", "Total Present", "Total Absent", "Attendance %"}, 0) {
                @Override
                public boolean isCellEditable(int r, int c) {
                    return false;
                }
            };
        }

        protected void generateAggregateReport() {
            Map<String, double[]> reportData;
            try {
                reportData = calculateReportData(aggStartDateChooser.getSelectedDate(), aggEndDateChooser.getSelectedDate(), getAggSubjectFilter()); // Use specific getter
            } catch (IllegalArgumentException e) {
                JOptionPane.showMessageDialog(this, e.getMessage(), "Input Error", JOptionPane.ERROR_MESSAGE);
                return;
            }
            aggregateReportModel.setRowCount(0);
            List<User> students = ExcelDataManager.getUsersByRole("Student");
            for (User student : students) {
                double[] counts = reportData.get(student.id);
                double present = counts[0], absent = counts[1], total = present + absent;
                double percentage = (total == 0) ? 100.0 : (present / total) * 100.0;
                aggregateReportModel.addRow(new Object[]{student.id, student.name, (int) total, (int) present, (int) absent, percentage});
            }
        }

        protected void generateAtRiskReport() {
            Map<String, double[]> reportData;
            double threshold;
            try {
                threshold = Double.parseDouble(atRiskThresholdField.getText());
                reportData = calculateReportData(atRiskStartDateChooser.getSelectedDate(), atRiskEndDateChooser.getSelectedDate(), getAtRiskSubjectFilter()); // Use specific getter
                // --- BUG FIX IS HERE ---
            } catch (NumberFormatException e) {
                JOptionPane.showMessageDialog(this, "Threshold must be a number (e.g., 75).", "Input Error", JOptionPane.ERROR_MESSAGE);
                return;
            } catch (IllegalArgumentException e) {
                JOptionPane.showMessageDialog(this, e.getMessage(), "Input Error", JOptionPane.ERROR_MESSAGE);
                return;
            }
            // --- END OF FIX ---
            atRiskReportModel.setRowCount(0);
            List<User> students = ExcelDataManager.getUsersByRole("Student");
            for (User student : students) {
                double[] counts = reportData.get(student.id);
                double present = counts[0], absent = counts[1], total = present + absent;
                double percentage = (total == 0) ? 100.0 : (present / total) * 100.0;
                if (total > 0 && percentage < threshold) {
                    atRiskReportModel.addRow(new Object[]{student.id, student.name, (int) total, (int) present, (int) absent, percentage});
                }
            }
            if (atRiskReportModel.getRowCount() == 0) {
                JOptionPane.showMessageDialog(this, "No students are below " + threshold + "% for this period.", "Report Complete", JOptionPane.INFORMATION_MESSAGE);
            }
        }

        protected void exportToCSV(JTable table, String reportName) {
            TableModel model = table.getModel();
            if (model.getRowCount() == 0) {
                JOptionPane.showMessageDialog(this, "No data to export.", "Export Error", JOptionPane.WARNING_MESSAGE);
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

    // --- NEW: Admin-Specific Report Panel (FIXED & REFACTORED) ---
    static class AdminReportsPanel extends BaseReportsPanel {

        // BUG FIX: Must NOT be final if they are initialized in methods called by super()
        private JComboBox<String> aggFilterDropdown;
        private JComboBox<String> atRiskFilterDropdown;

        AdminReportsPanel(MainFrame frame) {
            // BUG FIX: The super() constructor MUST be the very first line.
            // It will call the overridden create...ToolbarExtras() methods,
            // which will initialize the dropdowns.
            super(frame);
        }

        // Helper to create the model, called by the two methods below
        private DefaultComboBoxModel<String> createFilterModel() {
            DefaultComboBoxModel<String> filterModel = new DefaultComboBoxModel<>();
            filterModel.addElement("All Subjects");
            for (String s : SUBJECT_LIST) {
                filterModel.addElement(s);
            }
            return filterModel;
        }

        @Override
        protected JPanel createAggToolbarExtras() {
            // BUG FIX: Initialize the field here
            this.aggFilterDropdown = new JComboBox<>(createFilterModel());
            this.aggFilterDropdown.setFont(AppStyles.FONT_NORMAL);

            JPanel extraPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 5, 0));
            extraPanel.setOpaque(false);
            extraPanel.add(new JLabel("Filter Subject:"));
            extraPanel.add(this.aggFilterDropdown); // Add the initialized component
            return extraPanel;
        }

        @Override
        protected JPanel createAtRiskToolbarExtras() {
            // BUG FIX: Initialize the OTHER field here
            this.atRiskFilterDropdown = new JComboBox<>(createFilterModel());
            this.atRiskFilterDropdown.setFont(AppStyles.FONT_NORMAL);

            JPanel extraPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 5, 0));
            extraPanel.setOpaque(false);
            extraPanel.add(new JLabel("Filter Subject:"));
            extraPanel.add(this.atRiskFilterDropdown); // Add the initialized component
            return extraPanel;
        }

        @Override
        protected String getAggSubjectFilter() {
            return (String) aggFilterDropdown.getSelectedItem();
        }

        @Override
        protected String getAtRiskSubjectFilter() {
            return (String) atRiskFilterDropdown.getSelectedItem();
        }
    }

    // --- NEW: Staff-Specific Report Panel ---
    static class StaffReportsPanel extends BaseReportsPanel {

        private final String staffSubject;

        StaffReportsPanel(MainFrame frame) {
            super(frame);
            this.staffSubject = frame.getCurrentUser().subject;
        }

        private JPanel createLabelPanel() {
            JPanel extraPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 5, 0));
            extraPanel.setOpaque(false);
            JLabel subjectLabel = new JLabel("Subject: " + this.staffSubject);
            subjectLabel.setFont(AppStyles.FONT_BOLD);
            extraPanel.add(subjectLabel);
            return extraPanel;
        }

        @Override
        protected JPanel createAggToolbarExtras() {
            return createLabelPanel();
        }

        @Override
        protected JPanel createAtRiskToolbarExtras() {
            return createLabelPanel();
        }

        @Override
        protected String getAggSubjectFilter() {
            return this.staffSubject;
        }

        @Override
        protected String getAtRiskSubjectFilter() {
            return this.staffSubject;
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
                String name = nameField.getText().trim(), pass = new String(passField.getPassword()).trim();
                if (name.isEmpty() || pass.isEmpty()) {
                    JOptionPane.showMessageDialog(this, "Name and Password cannot be empty.", "Error", JOptionPane.ERROR_MESSAGE);
                    return;
                }
                String subject = isStaff ? (String) subjectSelector.getSelectedItem() : "";
                ExcelDataManager.addUser(name, pass, role, subject);
                refreshCallback.run();
                dispose();
            });
        }
    }

    // --- NEW: Custom Modern Date Picker Button (No external JAR needed) ---
    static class ModernDatePickerButton extends JButton {

        private LocalDate selectedDate;
        private final Frame owner;

        public ModernDatePickerButton(Frame owner) {
            this.owner = owner;
            setSelectedDate(LocalDate.now());
            this.setFont(AppStyles.FONT_NORMAL);
            this.setCursor(new Cursor(Cursor.HAND_CURSOR));
            this.setIcon(new ImageIcon(createCalendarIcon(AppStyles.SECONDARY_TEXT_COLOR))); // Simple icon
            this.setHorizontalTextPosition(SwingConstants.LEFT);
            this.setIconTextGap(8);
            this.addActionListener(e -> openDatePicker());
        }

        private void openDatePicker() {
            Point loc = this.getLocationOnScreen();
            ModernDatePickerDialog dialog = new ModernDatePickerDialog(owner, selectedDate);
            dialog.setLocation(loc.x, loc.y + this.getHeight());
            dialog.setVisible(true);
            if (dialog.getSelectedDate() != null) {
                setSelectedDate(dialog.getSelectedDate());
            }
        }

        public LocalDate getSelectedDate() {
            return selectedDate;
        }

        public void setSelectedDate(LocalDate date) {
            this.selectedDate = date;
            this.setText(selectedDate.format(GLOBAL_DATE_FORMATTER));
        }

        // Create a simple programmatic icon so we don't need image files
        private static java.awt.image.BufferedImage createCalendarIcon(Color color) {
            java.awt.image.BufferedImage img = new java.awt.image.BufferedImage(16, 16, java.awt.image.BufferedImage.TYPE_INT_ARGB);
            Graphics2D g2 = img.createGraphics();
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            g2.setColor(color);
            g2.drawRect(1, 3, 13, 12); // Box
            g2.fillRect(3, 1, 3, 4);  // Left hinge
            g2.fillRect(9, 1, 3, 4);  // Right hinge
            g2.fillRect(3, 7, 3, 2);  // Day box 1
            g2.fillRect(6, 7, 3, 2);  // Day box 2
            g2.fillRect(9, 7, 3, 2);  // Day box 3
            g2.dispose();
            return img;
        }
    }

    // --- NEW: Custom Calendar Dialog (Replaces JCalendar) ---
    static class ModernDatePickerDialog extends JDialog {

        private LocalDate selectedDate;
        private YearMonth currentYearMonth;
        private final JPanel calendarPanel;
        private final JLabel monthYearLabel;

        public ModernDatePickerDialog(Frame owner, LocalDate initialDate) {
            super(owner, true);
            this.selectedDate = initialDate;
            this.currentYearMonth = YearMonth.from(initialDate);
            setUndecorated(true); // Modern look
            setLayout(new BorderLayout(5, 5));
            getRootPane().setBorder(BorderFactory.createLineBorder(AppStyles.BORDER_COLOR, 1));
            ((JPanel) getContentPane()).setBorder(new EmptyBorder(5, 5, 5, 5));

            JPanel headerPanel = new JPanel(new BorderLayout());
            headerPanel.setBackground(Color.WHITE);
            JButton prevButton = createNavButton("‹");
            prevButton.addActionListener(e -> updateCalendar(currentYearMonth.minusMonths(1)));
            JButton nextButton = createNavButton("›");
            nextButton.addActionListener(e -> updateCalendar(currentYearMonth.plusMonths(1)));
            monthYearLabel = new JLabel("", SwingConstants.CENTER);
            monthYearLabel.setFont(AppStyles.FONT_BOLD.deriveFont(16f));
            headerPanel.add(prevButton, BorderLayout.WEST);
            headerPanel.add(monthYearLabel, BorderLayout.CENTER);
            headerPanel.add(nextButton, BorderLayout.EAST);
            add(headerPanel, BorderLayout.NORTH);

            calendarPanel = new JPanel(new GridLayout(0, 7, 2, 2)); // 7 columns
            calendarPanel.setBackground(Color.WHITE);
            add(calendarPanel, BorderLayout.CENTER);

            JButton todayButton = new JButton("Today");
            todayButton.setFont(AppStyles.FONT_BOLD);
            todayButton.addActionListener(e -> setSelectedDateAndClose(LocalDate.now()));
            add(todayButton, BorderLayout.SOUTH);

            updateCalendar(currentYearMonth);
            pack();
        }

        private void updateCalendar(YearMonth yearMonth) {
            currentYearMonth = yearMonth;
            calendarPanel.removeAll();
            monthYearLabel.setText(currentYearMonth.format(DateTimeFormatter.ofPattern("MMMM yyyy")));

            DayOfWeek firstDay = DayOfWeek.MONDAY;
            for (int i = 0; i < 7; i++) {
                String dayName = firstDay.plus(i).getDisplayName(TextStyle.SHORT, Locale.getDefault());
                JLabel headerLabel = new JLabel(dayName, SwingConstants.CENTER);
                headerLabel.setFont(AppStyles.FONT_BOLD);
                headerLabel.setForeground(AppStyles.SECONDARY_TEXT_COLOR);
                calendarPanel.add(headerLabel);
            }

            LocalDate firstOfMonth = currentYearMonth.atDay(1);
            int padding = (firstOfMonth.getDayOfWeek().getValue() - firstDay.getValue() + 7) % 7;
            for (int i = 0; i < padding; i++) {
                calendarPanel.add(new JLabel(""));
            }

            for (int day = 1; day <= currentYearMonth.lengthOfMonth(); day++) {
                calendarPanel.add(createDayButton(day));
            }
            pack(); // Resize dialog to fit new month
        }

        private JButton createDayButton(int day) {
            JButton dayButton = new JButton(String.valueOf(day));
            dayButton.setFont(AppStyles.FONT_NORMAL);
            dayButton.setFocusPainted(false);
            dayButton.setMargin(new Insets(8, 8, 8, 8));
            dayButton.setCursor(new Cursor(Cursor.HAND_CURSOR));
            LocalDate buttonDate = currentYearMonth.atDay(day);

            if (buttonDate.equals(LocalDate.now())) {
                dayButton.setForeground(AppStyles.PRIMARY_COLOR);
                dayButton.setFont(AppStyles.FONT_BOLD);
            }
            if (buttonDate.equals(selectedDate)) {
                dayButton.setBackground(AppStyles.PRIMARY_COLOR);
                dayButton.setForeground(Color.WHITE);
                dayButton.setFont(AppStyles.FONT_BOLD);
            } else {
                dayButton.setBackground(Color.WHITE);
                dayButton.setBorder(BorderFactory.createEmptyBorder(1, 1, 1, 1));
            }
            dayButton.addActionListener(e -> setSelectedDateAndClose(buttonDate));
            return dayButton;
        }

        private JButton createNavButton(String text) {
            JButton navButton = new JButton(text);
            navButton.setFont(AppStyles.FONT_BOLD.deriveFont(18f));
            navButton.setBorder(new EmptyBorder(5, 10, 5, 10));
            navButton.setContentAreaFilled(false);
            navButton.setCursor(new Cursor(Cursor.HAND_CURSOR));
            return navButton;
        }

        private void setSelectedDateAndClose(LocalDate date) {
            this.selectedDate = date;
            this.dispose();
        }

        public LocalDate getSelectedDate() {
            return selectedDate;
        }
    }

    // --- Excel Data Manager (ALL WRITE METHODS FIXED) ---
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

        // --- AUTHENTICATE METHOD (PATCHED) ---
        public static User authenticateUser(String username, String password) {
            try (FileInputStream fis = new FileInputStream(FILE_NAME); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheet(USERS_SHEET);
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) {
                        continue;
                    }
                    // FIX: Use helper to read any cell type (String/Number) and trim whitespace
                    String cellName = getStringValue(row.getCell(2), "").trim();
                    String cellPass = getStringValue(row.getCell(1), "").trim();

                    if (cellName.equals(username.trim()) && cellPass.equals(password.trim())) {
                        String id = getStringValue(row.getCell(0), "");
                        String role = getStringValue(row.getCell(3), "");
                        String subject = getStringValue(row.getCell(4), "");
                        // Pass the original (non-trimmed) inputs to the constructor as they were the ones validated
                        return new User(id, password, username, role, subject);
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
            return null;
        }

        // --- ADD USER (WRITE METHOD PATCHED) ---
        public static void addUser(String name, String password, String role, String subject) {
            String idPrefix = role.equalsIgnoreCase("Student") ? "STU" : "STAFF";
            int lastIdNum = getUsersByRole(role).stream()
                    .map(u -> u.id.replaceAll("[^0-9]", "")).filter(s -> !s.isEmpty())
                    .mapToInt(Integer::parseInt).max().orElse(0);
            String newId = String.format("%s%03d", idPrefix, lastIdNum + 1);

            Workbook workbook;
            try (FileInputStream fis = new FileInputStream(FILE_NAME)) {
                workbook = new XSSFWorkbook(fis);
            } catch (IOException e) {
                e.printStackTrace();
                return;
            }

            Sheet sheet = workbook.getSheet(USERS_SHEET);
            Row newRow = sheet.createRow(sheet.getLastRowNum() + 1);
            newRow.createCell(0).setCellValue(newId);
            newRow.createCell(1).setCellValue(password);
            newRow.createCell(2).setCellValue(name);
            newRow.createCell(3).setCellValue(role);
            newRow.createCell(4).setCellValue(subject != null ? subject : "");

            try (FileOutputStream fos = new FileOutputStream(FILE_NAME)) {
                workbook.write(fos);
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        // --- REMOVE USER (WRITE METHOD PATCHED) ---
        public static void removeUser(String id) {
            Workbook workbook;
            try (FileInputStream fis = new FileInputStream(FILE_NAME)) {
                workbook = new XSSFWorkbook(fis);
            } catch (IOException e) {
                e.printStackTrace();
                return;
            }

            Sheet sheet = workbook.getSheet(USERS_SHEET);
            int rowToRemove = -1;
            for (Row row : sheet) {
                if (row.getRowNum() == 0 || row.getCell(0) == null) {
                    continue;
                }
                if (getStringValue(row.getCell(0), "").equalsIgnoreCase(id)) {
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

                try (FileOutputStream fos = new FileOutputStream(FILE_NAME)) {
                    workbook.write(fos);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        // --- UPDATE PASSWORD (WRITE METHOD PATCHED) ---
        public static void updatePassword(String userId, String newPassword) {
            File file = new File(FILE_NAME);
            Workbook workbook;
            boolean updated = false;

            try (FileInputStream fis = new FileInputStream(file)) {
                workbook = new XSSFWorkbook(fis);
            } catch (IOException e) {
                e.printStackTrace();
                return;
            }

            Sheet sheet = workbook.getSheet(USERS_SHEET);
            for (Row row : sheet) {
                if (row.getRowNum() == 0 || row.getCell(0) == null) {
                    continue;
                }
                if (getStringValue(row.getCell(0), "").equalsIgnoreCase(userId)) {
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
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }

            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        // --- These READ-ONLY methods are fine as-is ---
        public static List<User> getUsersByRole(String role) {
            List<User> users = new ArrayList<>();
            try (FileInputStream fis = new FileInputStream(FILE_NAME); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheet(USERS_SHEET);
                for (Row row : sheet) {
                    if (row.getRowNum() == 0 || row.getCell(3) == null) {
                        continue;
                    }
                    if (row.getCell(3).getStringCellValue().equalsIgnoreCase(role)) {
                        String id = getStringValue(row.getCell(0), ""), name = getStringValue(row.getCell(2), "");
                        String subject = getStringValue(row.getCell(4), "");
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
                    if (getStringValue(row.getCell(0), "").equalsIgnoreCase(userId)) {
                        String name = getStringValue(row.getCell(2), ""), role = getStringValue(row.getCell(3), "");
                        String subject = getStringValue(row.getCell(4), "");
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
                    if (row.getRowNum() == 0 || row.getCell(0) == null || getStringValue(row.getCell(0), "").isEmpty()) {
                        continue;
                    }
                    String subject = getStringValue(row.getCell(3), "General");
                    records.add(new AttendanceRecord(getStringValue(row.getCell(0), ""), getStringValue(row.getCell(1), ""), getStringValue(row.getCell(2), ""), subject));
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
            return records;
        }

        // ... other read-only calculation methods are fine ...
        public static List<AttendanceRecord> getAllAttendanceForToday() {
            String today = LocalDate.now().format(GLOBAL_DATE_FORMATTER);
            return getAllAttendance().stream().filter(r -> r.date.equals(today)).collect(Collectors.toList());
        }

        public static Map<String, Double> getOverallAttendanceForToday() {
            return calculateStatsFromRecords(getAllAttendanceForToday());
        }

        public static Map<String, Double> getOverallAttendanceForToday(String subject) {
            List<AttendanceRecord> todayRecords = getAllAttendanceForToday().stream()
                    .filter(r -> r.subject.equals(subject)).collect(Collectors.toList());
            return calculateStatsFromRecords(todayRecords);
        }

        private static Map<String, Double> calculateStatsFromRecords(List<AttendanceRecord> records) {
            double present = 0, absent = 0;
            for (AttendanceRecord r : records) {
                if ("Present".equals(r.status)) {
                    present++;
                } else {
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
            return getAllAttendance().stream().filter(r -> r.studentId.equalsIgnoreCase(studentId))
                    .sorted((r1, r2) -> r2.date.compareTo(r1.date)).collect(Collectors.toList());
        }

        public static boolean hasAttendanceBeenMarked(String dateStr, String subject) {
            return getAllAttendance().stream().anyMatch(r -> r.date.equals(dateStr) && r.subject.equals(subject));
        }

        public static String getStudentStatusForDate(String studentId, String dateStr, String subject) {
            return getAllAttendance().stream()
                    .filter(r -> r.date.equals(dateStr) && r.studentId.equals(studentId) && r.subject.equals(subject))
                    .map(r -> r.status).findFirst().orElse("Absent");
        }

        // --- MARK ATTENDANCE (WRITE METHOD PATCHED) ---
        public static void markAttendance(List<AttendanceRecord> records, String dateStr, String subject) {
            Workbook workbook;
            try (FileInputStream fis = new FileInputStream(FILE_NAME)) {
                workbook = new XSSFWorkbook(fis);
            } catch (IOException e) {
                e.printStackTrace();
                return;
            }

            Sheet sheet = workbook.getSheet(ATTENDANCE_SHEET);
            List<Integer> rowsToRemove = new ArrayList<>();
            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                String rowDate = getStringValue(row.getCell(1), "");
                String rowSubject = getStringValue(row.getCell(3), "General");
                if (rowDate.equals(dateStr) && rowSubject.equals(subject)) {
                    rowsToRemove.add(row.getRowNum());
                }
            }
            rowsToRemove.sort(Comparator.reverseOrder());
            int numRowsRemoved = 0;
            for (int rowIndex : rowsToRemove) {
                Row row = sheet.getRow(rowIndex);
                if (row != null) {
                    sheet.removeRow(row);
                    numRowsRemoved++;
                }
            }
            if (numRowsRemoved > 0 && !rowsToRemove.isEmpty()) {
                int firstRowRemoved = rowsToRemove.get(rowsToRemove.size() - 1);
                if (firstRowRemoved + numRowsRemoved <= sheet.getLastRowNum()) {
                    sheet.shiftRows(firstRowRemoved + numRowsRemoved, sheet.getLastRowNum(), -numRowsRemoved, true, false);
                }
            }

            int lastRow = sheet.getLastRowNum();
            for (int i = 0; i < records.size(); i++) {
                Row newRow = sheet.createRow(lastRow + 1 + i);
                AttendanceRecord record = records.get(i);
                newRow.createCell(0).setCellValue(record.studentId);
                newRow.createCell(1).setCellValue(record.date);
                newRow.createCell(2).setCellValue(record.status);
                newRow.createCell(3).setCellValue(record.subject);
            }

            try (FileOutputStream fos = new FileOutputStream(FILE_NAME)) {
                workbook.write(fos);
            } catch (IOException e) {
                e.printStackTrace();
            } finally {
                try {
                    workbook.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        // --- UPDATE SINGLE RECORD (WRITE METHOD PATCHED) ---
        public static void updateSingleAttendanceRecord(String studentId, String date, String subject, String newStatus) {
            File file = new File(FILE_NAME);
            Workbook workbook;
            boolean updated = false;

            try (FileInputStream fis = new FileInputStream(file)) {
                workbook = new XSSFWorkbook(fis);
            } catch (Exception e) {
                e.printStackTrace();
                return;
            }

            Sheet sheet = workbook.getSheet(ATTENDANCE_SHEET);
            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                if (getStringValue(row.getCell(0), "").equals(studentId)
                        && getStringValue(row.getCell(1), "").equals(date)
                        && getStringValue(row.getCell(3), "General").equals(subject)) {
                    Cell statusCell = row.getCell(2);
                    if (statusCell == null) {
                        statusCell = row.createCell(2);
                    }
                    statusCell.setCellValue(newStatus);
                    updated = true;
                    break;
                }
            }

            if (updated) {
                try (FileOutputStream fos = new FileOutputStream(file)) {
                    workbook.write(fos);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }

            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        // --- This is the essential helper method ---
        private static String getStringValue(Cell cell, String defaultValue) {
            if (cell == null) {
                return defaultValue;
            }
            try {
                String val = cell.getStringCellValue();
                return (val == null || val.isEmpty()) ? defaultValue : val;
            } catch (Exception e) {
                try {
                    // Try to read as a number and convert to string
                    return String.valueOf((long) cell.getNumericCellValue());
                } catch (Exception e2) {
                    try {
                        // Handle formula cells that result in strings
                        return cell.getRichStringCellValue().getString();
                    } catch (Exception e3) {
                        try {
                            // Handle formula cells that result in numbers
                            return String.valueOf((long) cell.getNumericCellValue());
                        } catch (Exception e4) {
                            return defaultValue;
                        }
                    }
                }
            }
        }
    }
}
