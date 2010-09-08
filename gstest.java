import com.google.gdata.client.spreadsheet.FeedURLFactory;
import com.google.gdata.client.spreadsheet.SpreadsheetService;
import com.google.gdata.data.PlainTextConstruct;
import com.google.gdata.data.spreadsheet.*;
import com.google.gdata.util.ServiceException;

import java.io.IOException;
import java.net.URL;
import java.util.*;

public class gstest {

    gstest() {
    }

    String i2a(int i) {
        String pfx = i >= 26 ? i2a(i/26 - 1) : "";
        char ch = (char)('A' + (i % 26));
        return pfx + ch;
    }

    int a2i(String s) {
        s = s.toUpperCase();
        int value = 0;
        for (int i = 0; i < s.length(); ++i) {
            value = value * 26 + (s.charAt(i) - 'A');
        }
        return value;
    }

    SpreadsheetEntry getSpreadsheetEntry(SpreadsheetService service, String key)
            throws IOException, ServiceException {

        URL url = new URL("https://spreadsheets.google.com/feeds/spreadsheets/private/full");

        SpreadsheetFeed feed = service.getFeed(url, SpreadsheetFeed.class);
        List<SpreadsheetEntry> spreadsheets = feed.getEntries();
        String query = "key="+key;
        for (SpreadsheetEntry entry : spreadsheets) {
            String e_key = entry.getKey();
            URL link = new URL(entry.getSpreadsheetLink().getHref());
            if (e_key.equals(key) || link.getQuery().contains(query)) {
                return entry;
            }
        }
        return null;
    }

    WorksheetEntry getWorksheetEntry(SpreadsheetEntry se, String title)
            throws IOException, ServiceException {
        for (WorksheetEntry we : se.getWorksheets()) {
            if (we.getTitle().getPlainText().equals(title)) {
                return we;
            }
        }
        return null;
    }

    TableEntry getWorksheetTable(TableFeed tf, WorksheetEntry we)
            throws IOException, ServiceException {

        String title = we.getTitle().getPlainText();
        for(TableEntry table : tf.getEntries()) {
            if (table.getWorksheet().getName().equals(title)) {
                return table;
            }
        }
        return null;
    }

    TableEntry newWorksheetTable(TableFeed tf, WorksheetEntry we)
            throws IOException, ServiceException {
        String title = we.getTitle().getPlainText();
        TableEntry te = new TableEntry();
        te.setTitle(new PlainTextConstruct(title));
        te.setWorksheet(new Worksheet(title));
        te.setHeader(new Header(1));

        Data td = new Data();
        td.setNumberOfRows(0);
        td.setStartIndex(2);
        td.addColumn(new Column(i2a(0), "timestamp"));
        td.setInsertionMode(Data.InsertionMode.INSERT);
        te.setData(td);
        return tf.insert(te);
    }

    void updateTableColumns(WorksheetEntry we, TableEntry te, List<String> names)
            throws IOException, ServiceException {
        Set<String> exist = new HashSet<String>();
        int index = 0;

        List<Column> cols = te.getData().getColumns();
        for (Column c : cols) {
            exist.add(c.getName());
            int i = a2i(c.getIndex());
            if (index < i) {
                index = i;
            }
        }
        List<String> create = new ArrayList<String>();
        for (String key : names) {
            if (!exist.contains(key)) {
                create.add(key);
            }
        }
        if (!create.isEmpty()) {
            Data td = te.getData();
            int count = td.getColumns().size() + create.size();
            if (count > we.getColCount()) {
                we.setColCount(count);
                we.update();
            }
            for (String key : create) {
                td.addColumn(new Column(i2a(++index), key));
            }
            te.update();
        }
    }

    String getTableEntryId(String tid) {
        return tid.substring(tid.lastIndexOf("/") + 1);
    }

    void addWorksheetTableRows(SpreadsheetEntry se, WorksheetEntry we,
            List<List<String>> entries) throws ServiceException, IOException {

        List<String> names = entries.get(0);
        if (names.size() <= 0) {
            return;
        }

        TableFeed tf = se.getService().getFeed(FeedURLFactory.getDefault().getTableFeedUrl(se.getKey()),
                TableFeed.class);

        TableEntry te = getWorksheetTable(tf, we);
        if (te == null) {
            te = newWorksheetTable(tf, we);
        }

        updateTableColumns(we, te, names);

        RecordFeed rf = se.getService().getFeed(FeedURLFactory.getDefault().getRecordFeedUrl(
                se.getKey(), getTableEntryId(te.getId())), RecordFeed.class);

        for (int row = 1; row < entries.size(); ++row) {
            List<String> values = entries.get(row);
            assert names.size() == values.size();

            RecordEntry re = new RecordEntry();
            for (int i = 0; i < names.size(); ++i) {
                re.addField(new Field(null, names.get(i), values.get(i)));
            }
            rf.insert(re);
        }
    }

    void test(String u, String p, String key, String sheet,
              List<List<String>> entries) throws ServiceException, IOException {
        SpreadsheetService service = new SpreadsheetService("gstest");
        service.setUserCredentials(u, p);
        SpreadsheetEntry se = getSpreadsheetEntry(service, key);
        if (se == null) {
            System.err.println("ERROR: Can not find spreadsheet with key="+key);
            System.exit(-1);
        }
        WorksheetEntry we = getWorksheetEntry(se, sheet);
        if (we == null) {
            we = new WorksheetEntry();
            we.setTitle(new PlainTextConstruct(sheet));
            we.setRowCount(20);
            we.setColCount(5);
            we = se.getService().insert(se.getWorksheetFeedUrl(), we);
        }
        addWorksheetTableRows(se, we, entries);
    }

    public static void main(String[] args) {
        if (args.length < 4) {
            System.err.println("Usage: gstest username password spreadsheet_key worksheet_name\n" +
                    "\tusername - your gdocs username\n" +
                    "\tpassword - your gdocs password\n" +
                    "\tspreadsheet_key - key for existing spreadsheet\n" +
                    "\tworksheet_name - name of a worksheet in the spreadsheet, new or existing");
            System.exit(-1);
        }
        try {
            List<String> names = new ArrayList<String>();
            List<String> values = new ArrayList<String>();
            names.add("timestamp");
            values.add(new Date().toString());

            for (int i = 0; i < 50; ++i) {
                names.add("Name " + i);
                values.add("Value " + i);
            }

            List<List<String>> entries = new ArrayList<List<String>>();
            entries.add(names);
            entries.add(values);
            entries.add(values);
            entries.add(values);
            entries.add(values);
            entries.add(values);
            entries.add(values);
            entries.add(values);
            entries.add(values);
            entries.add(values);
            entries.add(values);

            System.err.println("Adding first set of rows...");
            new gstest().test(args[0], args[1], args[2], args[3], entries);

            for (int i = 50; i < 100; ++i) {
                names.add("Name " + i);
                values.add("Value " + i);
            }

            System.err.println("Adding second set of rows...");
            new gstest().test(args[0], args[1], args[2], args[3], entries);
        }
        catch (Throwable t) {
            t.printStackTrace();
        }
    }
}

