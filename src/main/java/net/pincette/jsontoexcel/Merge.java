package net.pincette.jsontoexcel;

import static java.lang.Integer.MAX_VALUE;
import static java.lang.System.exit;
import static java.time.Instant.parse;
import static java.util.logging.Logger.getGlobal;
import static java.util.stream.Collectors.toList;
import static java.util.stream.Stream.empty;
import static javax.json.Json.createParser;
import static javax.json.Json.createValue;
import static net.pincette.util.Collections.indexedStream;
import static net.pincette.util.Json.asNumber;
import static net.pincette.util.Json.asString;
import static net.pincette.util.Json.isInstant;
import static net.pincette.util.StreamUtil.rangeExclusive;
import static net.pincette.util.StreamUtil.stream;
import static net.pincette.util.StreamUtil.zip;
import static net.pincette.util.Util.matcherIterator;
import static net.pincette.util.Util.pathSearch;
import static net.pincette.util.Util.tryToDoWithRethrow;
import static net.pincette.util.When.when;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.function.BiFunction;
import java.util.function.Consumer;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;
import javax.json.JsonObject;
import javax.json.JsonValue;
import javax.json.stream.JsonParser;
import net.pincette.util.Json;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * With this class JSON can be merged with an Excel template, which should contain one sheet with a
 * row that has only bindings of the form "{field}" if the JSON is an array. If the JSON is an
 * object the template may have such bindings everywhere. The field may be a dot-separated path.
 *
 * @author Werner Donn\u00e9
 */
public class Merge {
  private static final Pattern BINDING = Pattern.compile("\\{([^\\{\\}]+)\\}");

  private Merge() {}

  private static SXSSFSheet addHeader(final SXSSFSheet out, final XSSFSheet in) {
    return zip(stream(in.getRow(0).iterator()), rangeExclusive(0, MAX_VALUE))
        .reduce(
            out.createRow(0),
            (r, pair) ->
                copyCell(
                    r,
                    pair.first,
                    pair.second,
                    cell -> cell.setCellValue(pair.first.getStringCellValue())),
            (r1, r2) -> r1)
        .getSheet();
  }

  private static SXSSFRow copyCell(
      final SXSSFRow row, final Cell cell, final int position, final Consumer<SXSSFCell> value) {
    final SXSSFCell copy = row.createCell(position, cell.getCellTypeEnum());
    final CellStyle style = row.getSheet().getWorkbook().createCellStyle();

    style.cloneStyleFrom(cell.getCellStyle());
    value.accept(copy);
    copy.setCellStyle(style);
    row.getSheet().setColumnWidth(position, cell.getSheet().getColumnWidth(position));

    return row;
  }

  private static SXSSFRow copyRow(
      final Row from, final SXSSFRow to, final JsonObject json, final CellStyle dateStyle) {
    return stream(from.iterator())
        .reduce(
            to,
            (r, c) ->
                copyCell(
                    r,
                    c,
                    c.getColumnIndex(),
                    cell ->
                        when(isBindingCell(c))
                            .run(
                                () ->
                                    setValue(
                                        cell, getValue(c.getStringCellValue(), json), dateStyle))
                            .orElse(() -> copyValue(c, cell))),
            (r1, r2) -> r1);
  }

  private static void copyValue(final Cell from, final SXSSFCell to) {
    switch (from.getCellTypeEnum()) {
      case BOOLEAN:
        to.setCellValue(from.getBooleanCellValue());
        break;
      case ERROR:
        to.setCellErrorValue(from.getErrorCellValue());
        break;
      case FORMULA:
        to.setCellValue(from.getCellFormula());
        break;
      case NUMERIC:
        to.setCellValue(from.getNumericCellValue());
        break;
      case STRING:
        to.setCellValue(from.getStringCellValue());
        break;
      default:
        break;
    }
  }

  private static CellStyle createDateStyle(final SXSSFWorkbook wb) {
    final CellStyle style = wb.createCellStyle();

    style.setDataFormat((short) 14);

    return style;
  }

  private static SXSSFSheet createSheet(final XSSFWorkbook in, final SXSSFWorkbook out) {
    final SXSSFSheet newSheet = out.createSheet();

    newSheet.untrackAllColumnsForAutoSizing();

    Optional.ofNullable(in.getSheetAt(0))
        .map(XSSFSheet::getPaneInformation)
        .filter(PaneInformation::isFreezePane)
        .ifPresent(freeze -> newSheet.createFreezePane(0, 1));

    return newSheet;
  }

  private static SXSSFRow fillRow(
      final SXSSFRow row,
      final JsonObject data,
      final List<String> bindings,
      final CellStyle dateStyle) {
    return indexedStream(bindings)
        .reduce(
            row,
            (r, pair) ->
                (SXSSFRow)
                    setValue(r.createCell(pair.second), getValue(pair.first, data), dateStyle)
                        .getRow(),
            (r1, r2) -> r1);
  }

  private static SXSSFSheet generateRows(
      final SXSSFSheet sheet,
      final Stream<JsonObject> data,
      final List<String> rowData,
      final CellStyle dateStyle) {
    return data.filter(Objects::nonNull)
        .reduce(
            sheet,
            (s, o) -> fillRow(s.createRow(s.getLastRowNum() + 1), o, rowData, dateStyle).getSheet(),
            (s1, s2) -> s1);
  }

  private static Optional<Row> getBindingRow(final XSSFSheet sheet) {
    return stream(sheet.iterator()).filter(Merge::isBindingRow).findFirst();
  }

  private static Stream<String> getBindings(final String cellValue) {
    return Optional.of(BINDING.matcher(cellValue)).map(Merge::getMatches).orElse(empty());
  }

  private static List<String> getCells(final Row row) {
    return stream(row.iterator()).map(Cell::getStringCellValue).collect(toList());
  }

  private static Stream<String> getMatches(final Matcher matcher) {
    return stream(matcherIterator(matcher, m -> m.group(1)));
  }

  private static JsonValue getValue(final String value, final JsonObject json) {
    return reduceValue(
        value,
        getBindings(value)
            .map(binding -> pathSearch(json, binding).orElse(createValue("")))
            .toArray(JsonValue[]::new));
  }

  private static boolean isBindingCell(final Cell cell) {
    return BINDING.matcher(cell.getStringCellValue()).find();
  }

  private static boolean isBindingRow(final Row row) {
    return stream(row.iterator()).allMatch(Merge::isBindingCell);
  }

  public static void main(final String[] args) throws Exception {
    if (args.length != 3) {
      usage();
    }

    merge(
        createParser(new FileInputStream(args[0])),
        new FileInputStream(args[1]),
        new FileOutputStream(args[2]));
  }

  public static void merge(
      final JsonParser parser, final InputStream template, final OutputStream out) {
    if (parser.hasNext()) {
      switch (parser.next()) {
        case START_OBJECT:
          merge(parser.getObject(), template, out);
          break;
        case START_ARRAY:
          merge(parser.getArrayStream(), template, out);
          break;
        default:
          break;
      }
    }
  }

  public static void merge(
      final JsonObject json, final InputStream template, final OutputStream out) {
    merge(template, out, (wb, newWb) -> merge(json, wb, newWb));
  }

  public static void merge(
      final Stream<JsonValue> stream, final InputStream template, final OutputStream out) {
    merge(template, out, (wb, newWb) -> merge(stream, wb, newWb));
  }

  private static void merge(
      final InputStream template,
      final OutputStream out,
      final BiFunction<XSSFWorkbook, SXSSFWorkbook, SXSSFWorkbook> run) {
    tryToDoWithRethrow(
        SXSSFWorkbook::new,
        newWb ->
            tryToDoWithRethrow(
                () -> new XSSFWorkbook(template), wb -> run.apply(wb, newWb).write(out)));
  }

  private static SXSSFWorkbook merge(
      final JsonObject json, final XSSFWorkbook wb, final SXSSFWorkbook out) {
    final CellStyle dateStyle = createDateStyle(out);
    final SXSSFSheet newSheet = createSheet(wb, out);

    return stream(wb.getSheetAt(0).rowIterator())
        .reduce(
            newSheet,
            (s, r) -> copyRow(r, s.createRow(r.getRowNum()), json, dateStyle).getSheet(),
            (s1, s2) -> s1)
        .getWorkbook();
  }

  private static SXSSFWorkbook merge(
      final Stream<JsonValue> stream, final XSSFWorkbook wb, final SXSSFWorkbook out) {
    final CellStyle dateStyle = createDateStyle(out);
    final SXSSFSheet newSheet = createSheet(wb, out);

    return Optional.ofNullable(wb.getSheetAt(0))
        .flatMap(Merge::getBindingRow)
        .map(Merge::getCells)
        .map(
            cells ->
                generateRows(
                    addHeader(newSheet, wb.getSheetAt(0)),
                    stream.filter(Json::isObject).map(JsonValue::asJsonObject),
                    cells,
                    dateStyle))
        .map(sheet -> out)
        .orElse(out);
  }

  private static JsonValue reduceValue(final String value, final JsonValue[] values) {
    return createValue(values.length == 0 ? "" : replaceMatches(value, values));
  }

  private static String replaceMatches(final String value, final JsonValue[] values) {
    return Arrays.stream(values)
        .map(net.pincette.util.Json::toNative)
        .map(Object::toString)
        .reduce(value, (result, v) -> BINDING.matcher(result).replaceFirst(v), (r1, r2) -> r1);
  }

  private static SXSSFCell setValue(
      final SXSSFCell cell, final JsonValue value, final CellStyle dateStyle) {
    switch (value.getValueType()) {
      case FALSE:
        cell.setCellValue(false);
        break;
      case NUMBER:
        cell.setCellValue(asNumber(value).doubleValue());
        break;

      case STRING:
        if (isInstant(value)) {
          cell.setCellValue(new Date(parse(asString(value).getString()).toEpochMilli()));
          cell.setCellStyle(dateStyle);
        } else {
          cell.setCellValue(asString(value).getString());
        }

        break;

      case TRUE:
        cell.setCellValue(true);
        break;
      default:
        cell.setCellValue("");
    }

    return cell;
  }

  private static void usage() {
    getGlobal().severe("Usage: net.pincette.jsontoexcel.Merge json template excel");
    exit(1);
  }
}
