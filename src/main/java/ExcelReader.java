import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
// класс читает данные из Excell файла и возвращает массив из номера строки и списка содержимого строки по ячейкам
public class ExcelReader {

    private final HashMap<Integer, List<Object>> data = new HashMap<>();
    /*
В качестве параметра метод принимает полный путь до файла Excel. Затем мы вызываем метод loadWorkbook(),
который возвращает интерфейс Workbook. Этот интерфейс имеет несколько реализаций в зависимости от формата файла (xls, xlsx и т.п.).
В одном файле (книге) Excel имеется несколько страниц (листов). Эти листы мы обходим с помощью итератора,
который возвращает метод sheetIterator(). Обработка каждого листа происходит в методе processSheet().
 */
    public void read(String filename) throws ExceptiionReadExcellFile {
        try (Workbook workbook = loadWorkbook(filename)) {
            assert workbook != null;
            var sheetIterator = workbook.sheetIterator();
            while (sheetIterator.hasNext()) {
                Sheet sheet = sheetIterator.next();
                processSheet(sheet, data);
            }
        }
        catch (IOException exception) {
            throw new ExceptiionReadExcellFile();
        }
    }
    public void clear() {
        data.clear();
    }
/*
Метод loadWorkbook() анализирует расширение файла и в зависимости от него определяет формат файла.
Сначала мы создаём FileInputStream, связанный с исходным файлом, затем создаём на основе этого потока нужную реализацию
интерфейса Workbook. Если расширение отличается от тех, которые мы ожидаем, то кидаем исключение.
 */

    private Workbook loadWorkbook (String fileName) throws IOException {
        var extension = fileName.substring(fileName.lastIndexOf(".") + 1).toLowerCase();
        var file = new FileInputStream(fileName);
        return switch (extension) {
            case "xls" -> new HSSFWorkbook(file);
            case "xlsx" -> new XSSFWorkbook(file);
            default -> throw new RuntimeException("Unknown file extension - " + extension);
        };
    }

/*
Метод processSheet() обрабатывает один лист из Excel. Метод getSheetName() возвращает имя листа,
которое обычно в Excel пишется внизу. Далее создаём мапу data, куда будем складывать построчно данные из ячеек таблицы.
Ключ такой мапы – это номер строки, а значение – все ячейки данной таблицы.
Обход всех строк листа выполняем также через итератор.
 */
    private void processSheet(Sheet sheet, HashMap<Integer, List<Object>> data) {
        var iterator = sheet.rowIterator();
        for (var rowIndex = 0; iterator.hasNext(); rowIndex++) {
            var row = iterator.next();
            processRow(data, rowIndex, row);
        }
    }
// Метод processRow() просто вызывает в цикле метод processCell() для каждой ячейки.
    private void processRow(HashMap<Integer, List<Object>> data, int rowIndex, Row row) {
        data.put(rowIndex, new ArrayList<>());
        for (var cell : row) {
            processCell(cell, data.get(rowIndex));
        }
    }
/*
В методе processCell() для каждой ячейки мы вначале смотрим формат ячейки с помощью метода getCellType()
 Всего есть 4 значимых типа для формата ячейки: это строка, число, булевый (логический) тип и формула,
 в которой значение ячейки вычисляется динамически на основании других ячеек.
 Значение каждого типа мы получаем с помощью соответствующего метода. Для текстовых значений мы используем getStringCellValue().
 Для числового типа мы сначала проверяем формат с помощью метода DateUtil.isCellDateFormatted() на предмет наличия в ней даты.
 Если метод возвращает true, то интерпретируем значение ячейки как дату с помощью метода getLocalDateTimeCellValue(),
 иначе берём значение как число с помощью getNumericCellValue(). При этом используем NumberToTextConverter,
 который преобразует числа в текст. Если его не использовать, то даже целые числа будут иметь один десятичный знак после запятой.
 Формула, возвращаемая методом getCellFormula() содержит буквенно-числовые имена ячеек и выглядит примерно так: «A2+C2*2».
 */
    private void processCell(Cell cell, List<Object> dataRow) {
        switch (cell.getCellType()) {
            case STRING -> dataRow.add(cell.getStringCellValue());
            case NUMERIC -> {
                if (DateUtil.isCellDateFormatted(cell)) {
                    dataRow.add(cell.getLocalDateTimeCellValue());
                }
                else {
                    dataRow.add(NumberToTextConverter.toText(cell.getNumericCellValue()));
                }
            }
            case BOOLEAN -> dataRow.add(cell.getBooleanCellValue());
            case FORMULA -> dataRow.add(cell.getCellFormula());
            default -> dataRow.add(" ");
        }
    }

    public HashMap<Integer, List<Object>> getData() {
        return data;
    }
}
