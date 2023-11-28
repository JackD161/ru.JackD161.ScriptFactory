import javax.swing.*;
import java.awt.*;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class MainClassGUI {
    private final String version = """
            1.0 alpha
            Разработчик: Холопкин Юрий (JackD161)
            e-mail: holopkin_yurik@mail.ru
            tel: +7-951-827-85-67
            """;
    private final String specification = "Инструменты для генерации скриптов";
    private final String descriptionSchemas = """
            Размаркировка - генерирует скрипты удаления признака маркировки
            """;
    private final String errReadExcellFile = "Ошибка чтения файла с данными";
    private final String errRqFields = "Не заполнены обязательные поля для генерации скрипта";
    private final String selectedSchema = "Выбран тип скрипта ";
    private JFrame window;
    private JLabel srcFileLabel;
    private JTextField srcFile;
    private JLabel outFileLAbel;
    private JTextField outFile;
    private JButton reset;
    private JButton confirm;
    private JButton clipboard;
    private JTextArea outputField;
    private JTextArea logField;
    private final JPanel center;
    private JCheckBox clearMark;
public MainClassGUI() {
    initFrame();
    window.setBounds(300, 100, 930, 400);
    window.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
    window.setLayout(new BorderLayout());

    clearMark = new JCheckBox("Размаркировка");
    srcFileLabel = new JLabel("Файл источник информации");
    srcFile = new JTextField(30);
    outFileLAbel = new JLabel("Папка для сохранения скриптов");
    outFile = new JTextField(30);
    reset = new JButton("Сброс");
    confirm = new JButton("Генерация");
    clipboard = new JButton("В буффер");
    outputField = new JTextArea(10,40);
    logField = new JTextArea(5,40);

    JScrollPane scrollOutPane = new JScrollPane(outputField);
    JScrollPane scrollLogPane = new JScrollPane(logField);

    JMenuBar bar = new JMenuBar();
    JMenu file = new JMenu("Файл");
    JMenu help = new JMenu("Помощь");
    JMenuItem saveAs = new JMenuItem("Сохранить результат как...");
    JMenuItem exit = new JMenuItem("Выход");
    JMenuItem aboutIt = new JMenuItem("О программе");
    JMenuItem instruction = new JMenuItem("Описание работы");
    JMenuItem definition = new JMenuItem("Описание видов скриптов");
    file.add(saveAs);
    file.add(exit);
    help.add(aboutIt);
    help.add(instruction);
    help.add(definition);
    bar.add(file);
    bar.add(help);

    JPanel footer = new JPanel();
    center = new JPanel();
    center.add(srcFileLabel);
    center.add(srcFile);
    center.add(outFileLAbel);
    center.add(outFile);

    JPanel right = new JPanel();
    JPanel left = new JPanel();

    left.setLayout(new BoxLayout(left, BoxLayout.PAGE_AXIS));
    right.setLayout(new BoxLayout(right, BoxLayout.PAGE_AXIS));

    left.add(clearMark);

    right.add(new JLabel("Вывод"));
    right.add(scrollOutPane);
    right.add(new JLabel("Log"));
    right.add(scrollLogPane);

    footer.add(confirm);
    footer.add(reset);
    footer.add(clipboard);
    footer.setLayout(new FlowLayout());

    window.getContentPane().add(BorderLayout.PAGE_START, bar);
    window.getContentPane().add(BorderLayout.PAGE_END, footer);
    window.getContentPane().add(BorderLayout.CENTER, center);
    window.getContentPane().add(BorderLayout.LINE_END, right);
    window.getContentPane().add(BorderLayout.LINE_START, left);

    repaint();

    confirm.addActionListener(e -> {
        ExcelReader reader = new ExcelReader();
        try {
            reader.read(srcFile.getText());
            HashMap<Integer, List<Object>> data = reader.getData();
            outputField.append(String.valueOf(data.size()));
            for (Map.Entry<Integer, List<Object>> pair : data.entrySet()) {
                String code = (String) pair.getValue().get(0);
                outputField.append(pair.getValue().get(1) + "\n");
                new ClearMark(outFile.getText(), code);
                log(code);
            }
        }
     catch (ExceptiionReadExcellFile exception){
        JOptionPane.showMessageDialog(window, errReadExcellFile, "Ошибка", JOptionPane.ERROR_MESSAGE);
    }
    });
    reset.addActionListener(e -> {
        repaint();
    });

    aboutIt.addActionListener(e -> JOptionPane.showMessageDialog(window,"Версия генератора скриптов " + version));
    instruction.addActionListener(e -> JOptionPane.showMessageDialog(window, specification));
    definition.addActionListener(e -> JOptionPane.showMessageDialog(window,descriptionSchemas));
    exit.addActionListener(e -> System.exit(0));
    saveAs.addActionListener(e -> {
        if (!outFile.getText().isBlank()) {
            JOptionPane.showMessageDialog(window, "Файлы сохраняются автоматически при генерации скриптов");
        }
        else {
            JOptionPane.showMessageDialog(window, "Не задан путь сохранения файла", "Ошибка", JOptionPane.WARNING_MESSAGE);
        }
        });
        clipboard.addActionListener(e -> {
            String inputScanField = outputField.getText();
            StringSelection stringSelection = new StringSelection(inputScanField);
            Clipboard systemClipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
            systemClipboard.setContents(stringSelection, null);
        });
        confirm.addActionListener(e -> {

        });
}
    private void log(String message) {
        logField.append("\n" + message);
    }
    private void repaint() {
        window.setVisible(true);
        window.repaint();
    }
    private void initFrame() {
        window = new JFrame("Генератор скриптов");
    }

    public static void main(String[] args) {
        new MainClassGUI();
    }
}
